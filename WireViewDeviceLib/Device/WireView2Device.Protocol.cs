using System;
using System.Runtime.InteropServices;

namespace WireView2.Device
{
    public partial class WireView2Device
    {
        // Keep in sync with firmware DEVICE_STR_LEN
        public const int DEVICE_STR_LEN = 32;

        private enum UsbCmd : byte
        {
            CMD_WELCOME,
            CMD_READ_VENDOR_DATA,
            CMD_READ_UID,
            CMD_READ_DEVICE_DATA,
            CMD_READ_SENSOR_VALUES,
            CMD_READ_CONFIG,
            CMD_WRITE_CONFIG,
            CMD_READ_CALIBRATION,
            CMD_WRITE_CALIBRATION,
            CMD_RSVD1,
            CMD_RSVD2,
            CMD_RSVD3,
            CMD_DEVICE_CMD,
            CMD_READ_BUILD_INFO,
            CMD_CLEAR_FAULTS,
            CMD_RESET = 0xF0,
            CMD_BOOTLOADER = 0xF1,
            CMD_NVM_CONFIG = 0xF2,
            CMD_NOP = 0xFF
        }

        private enum SensorTs
        {
            SENSOR_TS1,
            SENSOR_TS2
        }

        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        private struct VendorDataStruct
        {
            public byte VendorId;
            public byte ProductId;
            public byte FwVersion;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        private struct BuildStruct
        {
            public VendorDataStruct VendorData;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = DEVICE_STR_LEN)]
            public string ProductName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = DEVICE_STR_LEN)]
            public string BuildInfo;
            public byte ProductNameLength;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        private struct PowerSensor
        {
            public short Voltage;
            public uint Current;
            public uint Power;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        private struct SensorStruct
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
            public short[] Ts; // 0.1 °C
            public ushort Vdd; // mV

            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 6)]
            public PowerSensor[] PowerReadings;

            public uint TotalPower; // mW
            public uint TotalCurrent; // mA
            public ushort AvgVoltage; // mV
            public HpwrCapability HpwrCapability; // 8-bit enum
            public ushort FaultStatus;
            public ushort FaultLog;
        }

        private enum HpwrCapability : byte
        {
            PSU_CAP_600W = 0,
            PSU_CAP_450W = 1,
            PSU_CAP_300W = 2,
            PSU_CAP_150W = 3
        }

        // ===== Device config (matches firmware) =====

        public enum FAULT : byte
        {
            FAULT_OTP_TCHIP,
            FAULT_OTP_TS,
            FAULT_OCP,
            FAULT_WIRE_OCP,
            FAULT_OPP,
            FAULT_CURRENT_IMBALANCE
        }

        public enum NVM_CMD : byte
        {
            NVM_CMD_NONE,
            NVM_CMD_LOAD,
            NVM_CMD_STORE,
            NVM_CMD_RESET,
            NVM_CMD_LOAD_CAL,
            NVM_CMD_STORE_CAL,
            NVM_CMD_LOAD_CAL_FACTORY,
            NVM_CMD_STORE_CAL_FACTORY
        }


        public enum AVG : byte
        {
            AVG_22MS,
            AVG_44MS,
            AVG_89MS,
            AVG_177MS,
            AVG_354MS,
            AVG_709MS,
            AVG_1417MS,
            AVG_2834MS,
            AVG_5668MS
        }

        [StructLayout(LayoutKind.Sequential, Pack = 4, CharSet = CharSet.Ansi)]
        public struct DeviceConfigStructV0
        {
            public ushort Crc;
            public byte Version;

            [MarshalAs(UnmanagedType.ByValArray, SizeConst = DEVICE_STR_LEN)]
            public byte[] FriendlyName;
            public ushort FaultBuzzerEnable;
            public ushort FaultSoftPowerEnable;
            public ushort FaultHardPowerEnable;
            public short TsFaultThreshold; // 0.1 °C
            public byte OcpFaultThreshold; // A
            public byte WireOcpFaultThreshold; // 0.1A
            public ushort OppFaultThreshold; // W
            public byte CurrentImbalanceFaultThreshold; // %
            public byte CurrentImbalanceFaultMinLoad; // A
            public byte ShutdownWaitTime; // seconds
            public AVG Average;
        }

    }
}