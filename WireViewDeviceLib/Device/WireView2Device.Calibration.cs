using System.Runtime.InteropServices;

namespace WireView2.Device;

public partial class WireView2Device
{
    public const int SENSOR_POWER_NUM = 6;
    public const int SENSOR_TS_NUM = 4;

    public enum CalibrationSlot : byte
    {
        Volatile = 0,
        User = 1,
        Factory = 2
    }

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct CalibrationValueStruct
    {
        public short Offset;
        public short GainOffset;
    }

    // Matches firmware CalibrationStruct (68 bytes)
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct CalibrationStruct
    {
        public ushort Crc;
        public byte Version;

        [MarshalAs(UnmanagedType.ByValArray, SizeConst = SENSOR_POWER_NUM)]
        public CalibrationValueStruct[] PowerReadingVoltage;

        [MarshalAs(UnmanagedType.ByValArray, SizeConst = SENSOR_POWER_NUM)]
        public CalibrationValueStruct[] PowerReadingCurrent;

        [MarshalAs(UnmanagedType.ByValArray, SizeConst = SENSOR_TS_NUM)]
        public CalibrationValueStruct[] Ts;

        public static int ByteSize => Marshal.SizeOf<CalibrationStruct>();

        public static CalibrationStruct CreateEmpty() => new CalibrationStruct
        {
            PowerReadingVoltage = new CalibrationValueStruct[SENSOR_POWER_NUM],
            PowerReadingCurrent = new CalibrationValueStruct[SENSOR_POWER_NUM],
            Ts = new CalibrationValueStruct[SENSOR_TS_NUM]
        };

        public byte[] ToBytes() => StructToBytes(this);

        public static CalibrationStruct FromBytes(byte[] bytes) => BytesToStruct<CalibrationStruct>(bytes);

        /// <summary>Compares all calibration values; the CRC is computed by the device on save, so it is ignored.</summary>
        public bool ValuesEqual(CalibrationStruct other)
        {
            return Version == other.Version
                && ArraysEqual(PowerReadingVoltage, other.PowerReadingVoltage)
                && ArraysEqual(PowerReadingCurrent, other.PowerReadingCurrent)
                && ArraysEqual(Ts, other.Ts);
        }

        private static bool ArraysEqual(CalibrationValueStruct[] a, CalibrationValueStruct[] b)
        {
            if (a.Length != b.Length) return false;
            for (int i = 0; i < a.Length; i++)
            {
                if (a[i].Offset != b[i].Offset || a[i].GainOffset != b[i].GainOffset) return false;
            }
            return true;
        }
    }

    /// <summary>
    /// Reads a calibration slot. Volatile is always answered; User/Factory are only answered by the
    /// firmware when the stored data has a valid CRC, so null for those means "not programmed".
    /// </summary>
    public CalibrationStruct? ReadCalibration(CalibrationSlot slot = CalibrationSlot.Volatile)
    {
        if (!Connected || _port == null) return null;

        byte[]? buf = SendData(new[] { (byte)UsbCmd.CMD_READ_CALIBRATION, (byte)slot }, CalibrationStruct.ByteSize);
        if (buf == null) return null;

        return CalibrationStruct.FromBytes(buf);
    }

    /// <summary>Writes the volatile calibration (RAM). Use NvmCmd(NVM_CMD_STORE_CAL*) to persist it.</summary>
    public bool WriteCalibration(CalibrationStruct cal)
    {
        if (!Connected || _port == null) return false;

        var payload = cal.ToBytes();
        const int maxPayloadPerFrame = 62;

        lock (_port)
        {
            _port!.Open();
            try
            {
                _port!.DiscardInBuffer();

                for (int offset = 0; offset < payload.Length; offset += maxPayloadPerFrame)
                {
                    int bytesToWrite = Math.Min(maxPayloadPerFrame, payload.Length - offset);

                    var frame = new byte[2 + bytesToWrite];
                    frame[0] = (byte)UsbCmd.CMD_WRITE_CALIBRATION;
                    frame[1] = (byte)offset;
                    Buffer.BlockCopy(payload, offset, frame, 2, bytesToWrite);

                    _port!.Write(frame, 0, frame.Length);
                }
            }
            finally
            {
                _port!.Close();
            }
        }

        return true;
    }
}
