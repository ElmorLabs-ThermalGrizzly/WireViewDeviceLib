using System.IO.Ports;
using System.Runtime.InteropServices;

namespace WireView2.Device
{
    // Use SharedSerialPort instead of SerialPort
    using SerialPort = SharedSerialPort;

    public partial class WireViewBasicDevice : IWireViewDevice, IDisposable
    {
        private const string WelcomeMessage = "Thermal Grizzly WireView";

        private readonly string _portName;
        private readonly int _baud;
        private SerialPort? _port;

        private enum UsbCmd : byte
        {
            CMD_WELCOME,
            CMD_READ_VENDOR_DATA
        }

        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        private struct VendorDataStruct
        {
            public byte VendorId;
            public byte ProductId;
            public byte FwVersion;
        }

        public event EventHandler<DeviceData>? DataUpdated;
        public event EventHandler<bool>? ConnectionChanged;

        public bool Connected { get; private set; }
        public string DeviceName { get; private set; } = string.Empty;
        public string HardwareRevision { get; private set; } = string.Empty;
        public string FirmwareVersion { get; private set; } = string.Empty;
        public string UniqueId { get; private set; } = string.Empty;

        public string PortName => _portName;

        public int VendorId { get; private set; }
        public int ProductId { get; private set; }
        public int FirmwareId { get; private set; }

        public WireViewBasicDevice(string portName, int baud = 115200)
        {
            _portName = portName;
            _baud = baud;
        }

        public static List<WireViewBasicDevice> FindDevices(int baud = 115200)
        {
            var devices = new List<WireViewBasicDevice>();

            foreach (var port in Stm32PortFinder.FindMatchingComPorts())
            {
                var device = new WireViewBasicDevice(port, baud);
                try
                {
                    device.Connect();
                    if (device.Connected)
                    {
                        devices.Add(device);
                    }
                    else
                    {
                        device.Dispose();
                    }
                }
                catch
                {
                    device.Dispose();
                }
            }

            return devices;
        }

        public void Connect()
        {
            if (Connected) return;

            _port = new SerialPort(_portName, _baud, Parity.None, 8, StopBits.One);
            _port.ReadTimeout = 100;
            _port.WriteTimeout = 100;

            // First try to read welcome message without sending command
            var welcomeMessage = ReadArbitraryLengthString();

            if (string.CompareOrdinal(welcomeMessage, 0, WelcomeMessage, 0, WelcomeMessage.Length) == 0)
            {

                var vd = ReadVendorData();
                if (vd != null)
                {
                    VendorId = vd.Value.VendorId;
                    ProductId = vd.Value.ProductId;
                    FirmwareId = vd.Value.FwVersion;

                    HardwareRevision = $"{vd.Value.VendorId:X2}{vd.Value.ProductId:X2}";
                    FirmwareVersion = vd.Value.FwVersion.ToString();

                    Connected = true;
                    ConnectionChanged?.Invoke(this, true);

                    return;
                }
            }

            Connected = false;

        }

        public void Disconnect()
        {
            var wasConnected = Connected;

            Connected = false;

            CleanupPort();

            if (wasConnected)
            {
                ConnectionChanged?.Invoke(this, false);
            }
        }

        private VendorDataStruct? ReadVendorData()
        {
            if (_port == null)
            {
                return null;
            }

            var size = Marshal.SizeOf<VendorDataStruct>();
            byte[]? buf = SendCmd(UsbCmd.CMD_READ_VENDOR_DATA, size);
            return buf == null ? null : BytesToStruct<VendorDataStruct>(buf);
        }

        private byte[]? SendCmd(UsbCmd cmd, int responseSize = 0, bool rts = false)
        {
            return SendData(new[] { (byte)cmd }, responseSize, rts);
        }

        private byte[]? SendData(byte[] data, int responseSize = 0, bool rts = false)
        {
            if (_port == null)
            {
                return null;
            }

            byte[]? buf = null;
            lock (_port)
            {
                _port.Open();
                try
                {
                    _port.DiscardInBuffer();
                    if (rts)
                    {
                        _port.RtsEnable = true;
                    }

                    if (data.Length > 0)
                    {
                        _port.Write(data, 0, data.Length);
                    }

                    if (responseSize > 0)
                    {
                        buf = ReadExact(responseSize);
                    }

                    if (rts)
                    {
                        _port.RtsEnable = false;
                    }
                }
                finally
                {
                    _port.Close();
                }
            }

            return buf;
        }

        private byte[]? ReadExact(int size)
        {
            if (_port == null)
            {
                return null;
            }

            var buf = new byte[size];
            var offset = 0;
            const int timeout = 1000;
            var start = Environment.TickCount64;

            while (offset < size && Environment.TickCount64 - start < timeout)
            {
                if (_port.BytesToRead > 0)
                {
                    offset += _port.Read(buf, offset, size - offset);
                }
            }

            return offset == size ? buf : null;
        }

        private string? ReadArbitraryLengthString(bool rts = true, int timeout = 100, int maxLength = 64)
        {
            if (_port == null)
            {
                return null;
            }

            var offset = 0;
            byte[]? buf = new byte[maxLength];

            lock (_port)
            {
                _port.Open();
                try
                {
                    _port.DiscardInBuffer();
                    if (rts)
                    {
                        _port.RtsEnable = true;
                    }

                    var start = Environment.TickCount64;

                    while (offset < maxLength && Environment.TickCount64 - start < timeout)
                    {
                        if (_port.BytesToRead > 0)
                        {
                            offset += _port.Read(buf, offset, _port.BytesToRead);
                        }
                    }

                    if (rts)
                    {
                        _port.RtsEnable = false;
                    }
                }
                finally
                {
                    _port.Close();
                }
            }

            return System.Text.Encoding.ASCII.GetString(buf, 0, offset).TrimEnd('\0');
        }

        private static T BytesToStruct<T>(byte[] bytes) where T : struct
        {
            var handle = GCHandle.Alloc(bytes, GCHandleType.Pinned);
            try
            {
                return Marshal.PtrToStructure<T>(handle.AddrOfPinnedObject());
            }
            finally
            {
                handle.Free();
            }
        }

        private void CleanupPort()
        {
            if (_port == null)
            {
                return;
            }

            try
            {
                _port.Close();
            }
            catch
            {
            }

            try
            {
                _port.Dispose();
            }
            catch
            {
            }

            _port = null;
        }

        public void Dispose()
        {
            DataUpdated = null;
            Disconnect();
        }

    }
}
