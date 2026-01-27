using System;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Management;

namespace WireView2.Device
{
    // Use SharedSerialPort instead of SerialPort
    using SerialPort = SharedSerialPort;

    public static class Stm32PortFinder
    {
        public static List<string> FindMatchingComPorts()
        {
            var ports = new List<string>();

            if (System.OperatingSystem.IsWindows())
            {
                using var searcher = new ManagementObjectSearcher(
                    "SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%(COM%)%'");

                foreach (var obj in searcher.Get().Cast<ManagementObject>())
                {
                    var pnpId = obj["PNPDeviceID"] as string ?? string.Empty;
                    if (pnpId.StartsWith(@"USB\VID_0483&PID_5740", StringComparison.OrdinalIgnoreCase))
                    {
                        var name = obj["Name"] as string;
                        if (name == null) continue;

                        var start = name.LastIndexOf("(COM", StringComparison.OrdinalIgnoreCase);
                        var end = name.LastIndexOf(')');
                        if (start >= 0 && end > start)
                        {
                            var comPort = name.Substring(start + 1, end - start - 1);
                            ports.Add(comPort);
                        }
                    }
                }
            }
            else
            {
                // For non-Windows systems, we can implement other methods if needed
                // For now, return empty list
            }

            return ports;
        }
    }

    public partial class WireViewPro2Device : IWireViewDevice, IDisposable
    {
        public const string WelcomeMessage = "Thermal Grizzly WireView Pro II";
        private readonly string _portName;
        private readonly int _baud;
        private SerialPort? _port;
        private CancellationTokenSource? _cts;
        private Task? _worker;

        public event EventHandler<DeviceData>? DataUpdated;
        public event EventHandler<bool>? ConnectionChanged;

        public bool Connected { get; private set; }
        public string DeviceName => "WireView Pro II";
        public string HardwareRevision { get; private set; } = string.Empty;
        public string FirmwareVersion { get; private set; } = string.Empty;
        public string UniqueId { get; private set; } = string.Empty;

        private int _pollIntervalMs = 1000;
        public int PollIntervalMs
        {
            get => _pollIntervalMs;
            set => _pollIntervalMs = Math.Max(100, Math.Min(5000, value));
        }

        public WireViewPro2Device(string portName, int baud = 115200)
        {
            _portName = portName;
            _baud = baud;
        }

        public void Connect()
        {
            if (Connected) return;

            _port = new SerialPort(_portName, _baud, Parity.None, 8, StopBits.One);
            _port.ReadTimeout = 1000;
            _port.WriteTimeout = 1000;

            // First try to read welcome message without sending command
            if (!ReadWelcomeMessage(false))
            {
                Connected = false;
                return;
            }

            var vd = ReadVendorData();
            if (vd != null && vd.Value.VendorId == 0xEF && vd.Value.ProductId == 0x05)
            {
                HardwareRevision = $"{vd.Value.VendorId:X2}{vd.Value.ProductId:X2}";
                FirmwareVersion = vd.Value.FwVersion.ToString();

                UniqueId = ReadUid() ?? string.Empty;

                // Enable display updates just in case
                ScreenCmd(SCREEN_CMD.SCREEN_RESUME_UPDATES);

                Connected = true;
                ConnectionChanged?.Invoke(this, true);
            }
            else
            {
                Connected = false;
            }
            _port.RtsEnable = false;
            _port.Close();

            if(Connected)
            {
                _cts = new CancellationTokenSource();
                _worker = Task.Run(() => PollLoop(_cts.Token));
            }
        }

        public void Disconnect()
        {
            if (!Connected) return;

            try
            {
                _cts?.Cancel();
                _worker?.Wait(1000);
            }
            catch { }

            Connected = false;

            HardwareRevision = string.Empty;
            FirmwareVersion = string.Empty;
            UniqueId = string.Empty;

            ConnectionChanged?.Invoke(this, false);
        }

        public string? ReadBuildString()
        {
            if (!Connected || _port == null) return null;

            var size = Marshal.SizeOf<BuildStruct>();
            byte[]? buf = null;

            lock (_port) { 
                _port!.Open();
                _port!.DiscardInBuffer();
                _port.Write(new byte[] { (byte)UsbCmd.CMD_READ_BUILD_INFO }, 0, 1);

                buf = ReadExact(size);
                _port!.Close();
            }

            if (buf == null) return null;
            BuildStruct buildStruct = BytesToStruct<BuildStruct>(buf);

            return buildStruct.BuildInfo;
        }

        public Task EnterBootloaderAsync()
        {
            if (!Connected || _port == null) return Task.CompletedTask;

            return Task.Run(() =>
            {
                try
                {
                    lock (_port)
                    {
                        _port!.Open();
                        _port!.DiscardInBuffer();
                        _port.Write(new byte[] { (byte)UsbCmd.CMD_BOOTLOADER }, 0, 1);
                        _port!.Close();
                    }
                    Thread.Sleep(50);
                }
                catch
                {
                }
                finally
                {
                    try { Disconnect(); } catch { }
                }
            });
        }

        public DeviceConfigStruct? ReadConfig()
        {
            if (!Connected || _port == null) return null;

            var size = Marshal.SizeOf<DeviceConfigStruct>();
            byte[]? buf = null;

            lock (_port)
            {
                _port!.Open();
                _port!.DiscardInBuffer();
                _port.Write(new byte[] { (byte)UsbCmd.CMD_READ_CONFIG }, 0, 1);

                buf = ReadExact(size);
                _port!.Close();
            }

            if (buf == null) return null;
            return BytesToStruct<DeviceConfigStruct>(buf);
        }

        public void WriteConfig(DeviceConfigStruct config)
        {
            if (!Connected || _port == null) return;

            var payload = StructToBytes(config);

            var frame = new byte[64];
            frame[0] = (byte)UsbCmd.CMD_WRITE_CONFIG;

            lock (_port)
            {
                _port!.Open();
                _port!.DiscardInBuffer();

                const int maxPayloadPerFrame = 62;

                for (int offset = 0; offset < payload.Length; offset += maxPayloadPerFrame)
                {
                    int bytesToWrite = Math.Min(maxPayloadPerFrame, payload.Length - offset);

                    frame[1] = (byte)offset;
                    Buffer.BlockCopy(payload, offset, frame, 2, bytesToWrite);

                    _port!.Write(frame, 0, bytesToWrite + 2);
                }
                _port!.Close();
            }
        }

        public void NvmCmd(NVM_CMD cmd)
        {
            if (!Connected || _port == null) return;

            lock (_port)
            {
                _port!.Open();
                _port!.DiscardInBuffer();
                _port!.Write(new[] { (byte)UsbCmd.CMD_NVM_CONFIG, (byte)0x55, (byte)0xAA, (byte)0x55, (byte)0xAA, (byte)cmd }, 0, 6);
                _port!.Close();
            }
        }

        public void ScreenCmd(SCREEN_CMD cmd)
        {
            if (!Connected || _port == null) return;

            lock (_port)
            {
                _port!.Open();
                _port!.DiscardInBuffer();
                _port!.Write(new[] { (byte)UsbCmd.CMD_SCREEN_CHANGE, (byte)cmd }, 0, 2);
                _port!.Close();
            }
        }

        public void ClearFaults(int faultStatusMask = 0xFFFF, int faultLogMask = 0xFFFF)
        {
            if (!Connected || _port == null) return;

            lock (_port)
            {
                _port!.Open();
                _port!.DiscardInBuffer();
                _port!.Write(new[] { (byte)UsbCmd.CMD_CLEAR_FAULTS, (byte)(faultStatusMask & 0xFF), (byte)((faultStatusMask >> 8) & 0xFF), (byte)(faultLogMask & 0xFF), (byte)((faultLogMask >> 8) & 0xFF) }, 0, 5);
                _port!.Close();
            }
        }

        private void PollLoop(CancellationToken ct)
        {
            try
            {
                while (!ct.IsCancellationRequested)
                {
                    var sensors = ReadSensorValues();
                    if (sensors != null)
            {
                        var d = MapSensorStruct(sensors.Value);
                        DataUpdated?.Invoke(this, d);
                    }
                    Thread.Sleep(_pollIntervalMs);
                }
            }
            catch (Exception)
                    {
                Disconnect();
            }
        }

        private bool ReadWelcomeMessage(bool sendCmd = false)
        {
            if (_port == null) return false;

            var size = WelcomeMessage.Length + 1;

            byte[]? buf = null;
            lock (_port)
            {
                _port!.Open();
                _port!.RtsEnable = true;

                if (sendCmd)
                {
                    _port!.DiscardInBuffer();
                    _port!.Write(new byte[] { (byte)UsbCmd.CMD_WELCOME }, 0, 1);
                }

                buf = ReadExact(size);
                _port.RtsEnable = false;
                _port!.Close();
            }

            if (buf == null) return false;
            return System.Text.Encoding.ASCII.GetString(buf, 0, size).TrimEnd('\0').CompareTo(WelcomeMessage) == 0;
        }

        private VendorDataStruct? ReadVendorData()
        {
            if (_port == null) return null;

            var size = Marshal.SizeOf<VendorDataStruct>();
            byte[]? buf;
            lock (_port)
            {
                _port!.Open();
                _port!.DiscardInBuffer();
                _port!.Write(new byte[] { (byte)UsbCmd.CMD_READ_VENDOR_DATA }, 0, 1);
                buf = ReadExact(size);
                _port!.Close();
            }

            if (buf == null) return null;
            return BytesToStruct<VendorDataStruct>(buf);
        }

        private string? ReadUid()
        {
            if (_port == null) return null;

            const int uidBytes = 12;
            byte[]? buf = null;

            lock (_port)
            {
                _port!.Open();
                _port!.DiscardInBuffer();
                _port!.Write(new byte[] { (byte)UsbCmd.CMD_READ_UID }, 0, 1);
                buf = ReadExact(uidBytes);
                _port!.Close();
            }

            if (buf == null) return null;
            return BitConverter.ToString(buf).Replace("-", string.Empty);
        }

        private SensorStruct? ReadSensorValues()
        {
            if(_port == null) return null;

            var size = Marshal.SizeOf<SensorStruct>();

            byte[]? buf = null;
            lock (_port)
            {
                _port!.Open();
                _port!.DiscardInBuffer();
                _port!.Write(new byte[] { (byte)UsbCmd.CMD_READ_SENSOR_VALUES }, 0, 1);
                buf = ReadExact(size);
                _port.Close();
            }

            if (buf == null) return null;
            return BytesToStruct<SensorStruct>(buf);
        }

        private DeviceData MapSensorStruct(SensorStruct ss)
        {
            var dd = new DeviceData
            {
                Connected = true,
                HardwareRevision = HardwareRevision,
                FirmwareVersion = FirmwareVersion,
                OnboardTempInC = ss.Ts[(int)SensorTs.SENSOR_TS_IN] / 10.0,
                OnboardTempOutC = ss.Ts[(int)SensorTs.SENSOR_TS_OUT] / 10.0,
                ExternalTemp1C = ss.Ts[(int)SensorTs.SENSOR_TS3] / 10.0,
                ExternalTemp2C = ss.Ts[(int)SensorTs.SENSOR_TS4] / 10.0,
                PsuCapabilityW = ss.HpwrCapability == HpwrCapability.PSU_CAP_600W ? 600 :
                                  ss.HpwrCapability == HpwrCapability.PSU_CAP_450W ? 450 :
                                  ss.HpwrCapability == HpwrCapability.PSU_CAP_300W ? 300 :
                                  ss.HpwrCapability == HpwrCapability.PSU_CAP_150W ? 150 : 0,

                // NEW
                FaultStatus = ss.FaultStatus,
                FaultLog = ss.FaultLog
            };

            for (int i = 0; i < 6; i++)
            {
                dd.PinVoltage[i] = ss.PowerReadings[i].Voltage / 1000.0;
                dd.PinCurrent[i] = ss.PowerReadings[i].Current / 1000.0;
            }

            return dd;
        }

        private byte[]? ReadExact(int size)
        {
            var buf = new byte[size];
            int offset = 0;
            int timeout = 1000;
            var start = Environment.TickCount;

            while (offset < size && Environment.TickCount - start < timeout)
            {
                if (_port!.BytesToRead > 0)
                {
                    offset += _port!.Read(buf, offset, size - offset);
                }
            }
            return offset == size ? buf : null;
        }

        public static T BytesToStruct<T>(byte[] bytes) where T : struct
        {
            var handle = GCHandle.Alloc(bytes, GCHandleType.Pinned);
            try { return Marshal.PtrToStructure<T>(handle.AddrOfPinnedObject()); }
            finally { handle.Free(); }
        }

        public static byte[] StructToBytes<T>(T value) where T : struct
        {
            int size = Marshal.SizeOf<T>();
            var bytes = new byte[size];

            nint p = Marshal.AllocHGlobal(size);
            try
            {
                Marshal.StructureToPtr(value, p, false);
                Marshal.Copy(p, bytes, 0, size);
                return bytes;
            }
            finally
            {
                Marshal.FreeHGlobal(p);
            }
        }

        public void Dispose() => Disconnect();

    }
}