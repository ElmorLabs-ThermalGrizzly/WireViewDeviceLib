using System.Runtime.InteropServices;

namespace WireView2.Device;

public static class DeviceLogParser
{
    // Firmware constants / layout assumptions
    private const int SENSOR_POWER_NUM = 6;
    private const int SENSOR_TS_NUM = 4;

    public enum ENTRY_TYPE : byte
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
    public static ENTRY_TYPE DecodeType(uint data) => (ENTRY_TYPE)(data & 0b11u);
    public static uint DecodeTimestamp30(uint data) => (data >> 2) & 0x3FFF_FFFFu;

    public static IReadOnlyList<DATALOGGER_Entry> Parse(ReadOnlySpan<byte> data)
    {
        int EntrySizeBytes = Marshal.SizeOf<DATALOGGER_Entry>();

        var results = new List<DATALOGGER_Entry>(capacity: Math.Max(1024, data.Length / EntrySizeBytes));
        if (data.Length < EntrySizeBytes)
            return results;

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
                default:
                case ENTRY_TYPE.ENTRY_TYPE_SYSTEM_TIME:
                    // Not implemented yet, go to next entry
                    offset += EntrySizeBytes;
                    continue;

                case ENTRY_TYPE.ENTRY_TYPE_POWER_ON:

                    break;

                case ENTRY_TYPE.ENTRY_TYPE_MCU_TICK:

                    // If timestamp is 0, skip entry
                    /*if (ts30 == 0)
                    {
                        offset += EntrySizeBytes;
                        continue;
                    }*/

                    break;
            }

            // Parse struct from bytes
            DATALOGGER_Entry entryStruct = WireViewPro2Device.BytesToStruct<DATALOGGER_Entry>(entry);

            if(entryStruct.HpwrSense > 3)
            {
                // Invalid value, skip entry
                offset += EntrySizeBytes;
                continue;
            }

            results.Add(entryStruct);
            firstEntryFound = true;
            offset += EntrySizeBytes;
        }

        return results;
    }

    private static uint ReadU32LE(ReadOnlySpan<byte> s, int o)
        => (uint)(s[o] | (s[o + 1] << 8) | (s[o + 2] << 16) | (s[o + 3] << 24));
}