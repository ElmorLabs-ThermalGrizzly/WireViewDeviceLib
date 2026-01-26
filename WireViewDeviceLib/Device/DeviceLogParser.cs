using System.Runtime.InteropServices;

namespace WireView2.Device;

public static class DeviceLogParser
{
    // Firmware constants / layout assumptions
    private const int SENSOR_POWER_NUM = 6;
    private const int SENSOR_TS_NUM = 4;

    private enum ENTRY_TYPE : byte
    {
        ENTRY_TYPE_MCU_TICK = 0x00,
        ENTRY_TYPE_SYSTEM_TIME = 0x01,
        ENTRY_TYPE_POWER_ON = 0x02,
        ENTRY_TYPE_EMPTY = 0x03
    }

    // Pack 1
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct DATALogger_PowerSensor
    {
        public byte Voltage; // 100mV
        public byte Current; // 100mA
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct DATALOGGER_Entry
    {
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

    public static IReadOnlyList<DATALOGGER_Entry> Parse(ReadOnlySpan<byte> data)
    {
        int EntrySizeBytes = Marshal.SizeOf<DATALOGGER_Entry>();

        var results = new List<DATALOGGER_Entry>(capacity: Math.Max(1024, data.Length / EntrySizeBytes));
        if (data.Length < EntrySizeBytes)
            return results;

        uint lastMcuTick30 = 0;
        DateTime mcuBaseUtc = DateTime.Parse("2026-01-01 00:00");

        bool firstEntryFound = false;
        int emptyCount = 0;

        for (int offset = 0; offset + EntrySizeBytes <= data.Length;)
        {
            var entry = data.Slice(offset, EntrySizeBytes).ToArray();
            uint rawData = ReadU32LE(entry, 0);
            var type = DecodeType(rawData);
            uint ts30 = DecodeTimestamp30(rawData);

            // If near page end, go to next page and skip nearby entries

            if (firstEntryFound && (offset & 0xFF) > (256 - EntrySizeBytes))
            {
                int remainingInPage = 256 - (offset & 0xFF);
                int skipBytesInNextPage = EntrySizeBytes - remainingInPage;
                offset += remainingInPage + skipBytesInNextPage;
                continue;
            }

            // Ignore entries explicitly marked empty/reserved
            if (type is ENTRY_TYPE.ENTRY_TYPE_EMPTY)
            {
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

            switch (type)
            {
                case ENTRY_TYPE.ENTRY_TYPE_SYSTEM_TIME:
                    offset++;
                    continue;

                case ENTRY_TYPE.ENTRY_TYPE_POWER_ON:

                    mcuBaseUtc.AddDays(1);
                    continue;

                case ENTRY_TYPE.ENTRY_TYPE_MCU_TICK:
                default:

                    // If timestamp is 0, skip entry
                    if (ts30 == 0)
                    {
                        offset += EntrySizeBytes;
                        continue;
                    }

                    // if timestamp is less than previous value, treat as power on
                    if (ts30 < lastMcuTick30)
                    {
                        mcuBaseUtc.AddDays(1);
                    }

                    // 30-bit wrap-safe delta
                    uint deltaTicks = (ts30 - lastMcuTick30) & 0x3FFF_FFFFu;
                    lastMcuTick30 = ts30;

                    // Each tick = 4ms
                    mcuBaseUtc = mcuBaseUtc.AddMilliseconds(deltaTicks * 4.0);
                    break;
            }

            // Parse struct from bytes
            DATALOGGER_Entry entryStruct = WireViewPro2Device.BytesToStruct<DATALOGGER_Entry>(entry);

            results.Add(entryStruct);
            firstEntryFound = true;
            offset += EntrySizeBytes;
        }

        return results;
    }

    private static uint ReadU32LE(ReadOnlySpan<byte> s, int o)
        => (uint)(s[o] | (s[o + 1] << 8) | (s[o + 2] << 16) | (s[o + 3] << 24));
}