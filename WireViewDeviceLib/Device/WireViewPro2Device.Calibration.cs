using System.Runtime.InteropServices;

namespace WireView2.Device;

public partial class WireViewPro2Device
{
    public const int SENSOR_POWER_NUM = 6;
    public const int SENSOR_TS_NUM = 4;

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct CalibrationValueStruct
    {
        public short Offset;
        public short GainOffset;
    }

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

        public byte[] ToBytes()
        {
            var size = Marshal.SizeOf<CalibrationStruct>();
            var arr = new byte[size];
            var handle = GCHandle.Alloc(arr, GCHandleType.Pinned);
            try
            {
                Marshal.StructureToPtr(this, handle.AddrOfPinnedObject(), false);
                return arr;
            }
            finally
            {
                handle.Free();
            }
        }

        public static CalibrationStruct FromBytes(byte[] bytes)
        {
            var handle = GCHandle.Alloc(bytes, GCHandleType.Pinned);
            try
            {
                return Marshal.PtrToStructure<CalibrationStruct>(handle.AddrOfPinnedObject());
            }
            finally
            {
                handle.Free();
            }
        }
    }

    /// <summary>Firmware expects 2 bytes: command + selector (0=volatile, 1=user, 2=factory).</summary>
    public CalibrationStruct? ReadCalibration(int type = 0)
    {
        if (!Connected || _port == null) return null;

        byte[]? buf;
        lock (_port)
        {
            _port!.Open();
            try
            {
                _port!.DiscardInBuffer();
                _port!.Write(new byte[] { (byte)UsbCmd.CMD_READ_CALIBRATION, (byte)type }, 0, 2);
                buf = ReadExact(CalibrationStruct.ByteSize);
            }
            finally
            {
                _port!.Close();
            }
        }

        if (buf == null) return null;
        return CalibrationStruct.FromBytes(buf);
    }

    public bool WriteCalibration(CalibrationStruct cal)
    {
        if (!Connected || _port == null) return false;

        var payload = cal.ToBytes();
        const int firstChunkLen = 62;

        lock (_port)
        {
            _port!.Open();
            try
            {
                _port!.DiscardInBuffer();

                var frame = new byte[2 + firstChunkLen];
                frame[0] = (byte)UsbCmd.CMD_WRITE_CALIBRATION;
                frame[1] = 0;
                Buffer.BlockCopy(payload, 0, frame, 2, firstChunkLen);
                _port!.Write(frame, 0, frame.Length);

                int remaining = CalibrationStruct.ByteSize - firstChunkLen;
                frame = new byte[2 + remaining];
                frame[0] = (byte)UsbCmd.CMD_WRITE_CALIBRATION;
                frame[1] = firstChunkLen;
                Buffer.BlockCopy(payload, firstChunkLen, frame, 2, remaining);
                _port!.Write(frame, 0, frame.Length);
            }
            finally
            {
                _port!.Close();
            }
        }

        return true;
    }
}
