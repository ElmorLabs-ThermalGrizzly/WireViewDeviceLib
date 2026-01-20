using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace WireView2.Device;

internal static class DeviceLogParser
{
    // Firmware constants / layout assumptions
    private const int SENSOR_POWER_NUM = 6;
    private const int SENSOR_TS_NUM = 4;

    private enum ENTRY_TYPE : byte
    {
        ENTRY_TYPE_MCU_TICK = 0x00,
        ENTRY_TYPE_SYSTEM_TIME = 0x01,
        ENTRY_TYPE_RESERVED = 0x02,
        ENTRY_TYPE_EMPTY = 0x03
    }

    // Pack 1
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    private struct DATALogger_PowerSensor {
        public byte Voltage; // 100mV
        public byte Current; // 100mA
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    private struct DATALOGGER_Entry {
        public uint Data; // Type:2 + Timestamp:30
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = SENSOR_TS_NUM)]
        public byte[] Ts; // Temperatures
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = SENSOR_POWER_NUM)]
        public DATALogger_PowerSensor[] PowerSensors;
        public byte HpwrSense;
    }

    // Decode helpers for the union bitfields
    private static ENTRY_TYPE DecodeType(uint data) => (ENTRY_TYPE)(data & 0b11u);
    private static uint DecodeTimestamp30(uint data) => (data >> 2) & 0x3FFF_FFFFu;

    public static IReadOnlyList<DeviceData> Parse(ReadOnlySpan<byte> data)
    {
        int EntrySizeBytes = Marshal.SizeOf<DATALOGGER_Entry>();

        var results = new List<DeviceData>(capacity: Math.Max(1024, data.Length / EntrySizeBytes));
        if (data.Length < EntrySizeBytes)
            return results;

        bool haveMcuBase = false;
        uint lastMcuTick30 = 0;
        DateTime mcuBaseUtc = DateTime.UtcNow;

        bool firstEntryFound = false;
        int emptyCount = 0;

        for (int offset = 0; offset + EntrySizeBytes <= data.Length;)
        {
            var entry = data.Slice(offset, EntrySizeBytes).ToArray();
            uint rawData = ReadU32LE(entry, 0);
            var type = DecodeType(rawData);
            uint ts30 = DecodeTimestamp30(rawData);

            // Ignore entries explicitly marked empty/reserved
            if (type is ENTRY_TYPE.ENTRY_TYPE_EMPTY or ENTRY_TYPE.ENTRY_TYPE_RESERVED) {
                offset++;
                if (firstEntryFound)
                {
                    emptyCount++;
                    if (emptyCount >= 32)
                    {
                        break; // Stop parsing after 32 consecutive empty entries once we've found data
                    }
                }
                continue;
            }

            DateTime timestampUtc;
            switch (type)
            {
                case ENTRY_TYPE.ENTRY_TYPE_SYSTEM_TIME:
                    // Timestamp is seconds since epoch
                    //timestampUtc = DateTime.UnixEpoch.AddSeconds(ts30);
                    // Also reset MCU base so subsequent MCU ticks don't "jump back" unexpectedly
                    //haveMcuBase = false;
                    continue;
                    break;

                case ENTRY_TYPE.ENTRY_TYPE_MCU_TICK:
                default:
                    // Timestamp is MCU tick in units of 4ms, stored as 30-bit counter.
                    // We keep a rolling base and advance by delta ticks.
                    if (!haveMcuBase)
                    {
                        lastMcuTick30 = ts30;
                        mcuBaseUtc = DateTime.UtcNow;
                        haveMcuBase = true;
                    }

                    // 30-bit wrap-safe delta
                    uint deltaTicks = (ts30 - lastMcuTick30) & 0x3FFF_FFFFu;
                    lastMcuTick30 = ts30;

                    // Each tick = 4ms
                    mcuBaseUtc = mcuBaseUtc.AddMilliseconds(deltaTicks * 4.0);
                    timestampUtc = mcuBaseUtc;
                    break;
            }

            // Parse struct from bytes
            DATALOGGER_Entry entryStruct = WireViewPro2Device.BytesToStruct<DATALOGGER_Entry>(entry);

            var d = new DeviceData
            {
                Timestamp = timestampUtc,
                Connected = true,
                HardwareRevision = string.Empty,
                FirmwareVersion = string.Empty,

                OnboardTempInC = entryStruct.Ts[0],
                OnboardTempOutC = entryStruct.Ts[1],
                ExternalTemp1C = 0,
                ExternalTemp2C = 0,
            };

            // Firmware: Voltage in 100mV, Current in 100mA
            // => V = raw / 10.0, A = raw / 10.0
            for (int i = 0; i < SENSOR_POWER_NUM; i++)
            {
                d.PinVoltage[i] = entryStruct.PowerSensors[i].Voltage / 10.0;
                d.PinCurrent[i] = entryStruct.PowerSensors[i].Current / 10.0;
            }

            results.Add(d);
            firstEntryFound = true;
            offset += EntrySizeBytes;
        }

        return results;
    }

    private static uint ReadU32LE(ReadOnlySpan<byte> s, int o)
        => (uint)(s[o] | (s[o + 1] << 8) | (s[o + 2] << 16) | (s[o + 3] << 24));
}