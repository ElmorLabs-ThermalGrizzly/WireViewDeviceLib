using System.IO.Ports;
using System.Runtime.InteropServices;

namespace WireView2.Device
{
    // Use SharedSerialPort instead of SerialPort
    using SerialPort = SharedSerialPort;

    public partial class WireView2Device : IWireViewDevice, IDisposable
    {
        public string WelcomeMessage { get; private set; }

        private int VendorId;
        private int ProductId;

        private static readonly int SensorStructSize = Marshal.SizeOf<SensorStruct>();
        private static readonly int SensorStructNoFanSize = Marshal.SizeOf<SensorStructNoFan>();

        // False for firmware that sends the compact sensor layout (SensorStructNoFan)
        protected virtual bool HasFanSensors => true;

        // Set on connect by probing the size of the sensor reply; overrides HasFanSensors when known.
        private bool? _hasFanSensorsDetected;
        private bool UseFanSensorLayout => _hasFanSensorsDetected ?? HasFanSensors;

        private readonly string _portName;
        private readonly int _baud;

        public string PortName => _portName;
        private SerialPort? _port;
        private CancellationTokenSource? _cts;
        private Task? _worker;

        public event EventHandler<DeviceData>? DataUpdated;
        public event EventHandler<bool>? ConnectionChanged;

        public bool Connected { get; private set; }
        public string DeviceName { get; private set; }
        public string HardwareRevision { get; private set; } = string.Empty;
        public string FirmwareVersion { get; private set; } = string.Empty;
        public string UniqueId { get; private set; } = string.Empty;
        public Guid Uid { get; private set; } = Guid.Empty;
        public DeviceData? LastDeviceData { get; private set; }
        public virtual PowerControlType PowerControl => PowerControlType.PowerButton;

        public int ConfigVersion { get; private set; }

        private int _pollIntervalMs = 1000;
        public int PollIntervalMs
        {
            get => _pollIntervalMs;
            set => _pollIntervalMs = Math.Max(10, Math.Min(5000, value));
        }

        public WireView2Device(string portName, int baud = 115200, string welcomeMessage = "Thermal Grizzly WireView II", string deviceName = "WireView II", int vendorId = 0xEF, int productId = 0x07)
        {
            _portName = portName;
            _baud = baud;
            WelcomeMessage = welcomeMessage;
            DeviceName = deviceName;
            VendorId = vendorId;
            ProductId = productId;
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
            if (vd != null && vd.Value.VendorId == VendorId && vd.Value.ProductId == ProductId)
            {
                HardwareRevision = $"{vd.Value.VendorId:X2}{vd.Value.ProductId:X2}";
                FirmwareVersion = vd.Value.FwVersion.ToString();

                // Get config version
                int? configVersion = ReadConfigVersion();
                if (configVersion == null)
                {
                    Connected = false;
                    return;
                }

                ConfigVersion = configVersion.Value;

                var uidBytes = ReadUidBytes();
                if (uidBytes != null)
                {
                    UniqueId = BitConverter.ToString(uidBytes).Replace("-", string.Empty);
                    var guidBytes = uidBytes;
                    Array.Resize(ref guidBytes, 16);
                    Uid = new Guid(guidBytes);
                }
                else
                {
                    UniqueId = string.Empty;
                    Uid = Guid.Empty;
                }

                DetectSensorLayout();

                Connected = true;
                ConnectionChanged?.Invoke(this, true);
            }
            else
            {
                Connected = false;
            }

            if (Connected)
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
            Uid = Guid.Empty;
            LastDeviceData = null;
            _hasFanSensorsDetected = null;

            ConnectionChanged?.Invoke(this, false);
        }

        public string? ReadBuildString()
        {
            if (!Connected || _port == null) return null;

            var size = Marshal.SizeOf<BuildStruct>();
            byte[]? buf = SendCmd(UsbCmd.CMD_READ_BUILD_INFO, size);

            if (buf == null) return null;
            BuildStruct buildStruct = BytesToStruct<BuildStruct>(buf);

            return buildStruct.BuildInfo;
        }

        public void EnterBootloader()
        {
            if (!Connected || _port == null) return;

            SendCmd(UsbCmd.CMD_BOOTLOADER);
            try { Disconnect(); } catch { }
        }

        public Task EnterBootloaderAsync() => Task.Run(EnterBootloader);

        public DeviceConfigStructV0? ReadConfig()
        {
            if (!Connected || _port == null) return null;

            var size = 0;

            if (ConfigVersion == 0)
            {
                size = Marshal.SizeOf<DeviceConfigStructV0>();
            }
            else
            {
                return null;
            }

            byte[]? buf = SendCmd(UsbCmd.CMD_READ_CONFIG, size);

            if (buf == null) return null;

            if (ConfigVersion == 0)
            {
                var _s = BytesToStruct<DeviceConfigStructV0>(buf);
                return _s;
            }
            else
            {
                return null;
            }
        }

        public void WriteConfig(DeviceConfigStructV0 config)
        {
            if (!Connected || _port == null) return;

            var payload = new byte[0];

            if (ConfigVersion == 0)
            {
                payload = StructToBytes(config);
            }
            else
            {
                return;
            }

            var frame = new byte[64];
            frame[0] = (byte)UsbCmd.CMD_WRITE_CONFIG;

            lock (_port)
            {
                _port!.Open();
                try
                {
                    _port!.DiscardInBuffer();

                    const int maxPayloadPerFrame = 62;

                    for (int offset = 0; offset < payload.Length && offset <= 255; offset += maxPayloadPerFrame)
                    {
                        int bytesToWrite = Math.Min(maxPayloadPerFrame, payload.Length - offset);

                        frame[1] = (byte)offset;
                        Buffer.BlockCopy(payload, offset, frame, 2, bytesToWrite);

                        _port!.Write(frame, 0, bytesToWrite + 2);
                    }
                }
                finally
                {
                    _port!.Close();
                }
            }
        }

        public bool NvmCmd(NVM_CMD cmd)
        {
            if (!Connected || _port == null) return false;
            SendData(new[] { (byte)UsbCmd.CMD_NVM_CONFIG, (byte)0x55, (byte)0xAA, (byte)0x55, (byte)0xAA, (byte)cmd }, 0);
            return true;
        }

        public void ClearFaults(int faultStatusMask = 0xFFFF, int faultLogMask = 0xFFFF)
        {
            if (!Connected || _port == null) return;
            SendData(new[] { (byte)UsbCmd.CMD_CLEAR_FAULTS, (byte)(faultStatusMask & 0xFF), (byte)((faultStatusMask >> 8) & 0xFF), (byte)(faultLogMask & 0xFF), (byte)((faultLogMask >> 8) & 0xFF) }, 0);
        }

        private async Task PollLoop(CancellationToken ct)
        {
            try
            {
                while (!ct.IsCancellationRequested)
                {
                    var startedAt = Environment.TickCount64;

                    var sensors = ReadSensorValues();
                    if (sensors == null)
                    {
                        // No response (e.g. device unplugged): treat as disconnected.
                        Disconnect();
                        return;
                    }

                    var d = MapSensorStruct(sensors.Value);
                    LastDeviceData = d;
                    DataUpdated?.Invoke(this, d);

                    // Subtract the read time so the cadence stays fixed rather than drifting by it.
                    var elapsed = Environment.TickCount64 - startedAt;
                    await Task.Delay((int)Math.Max(0, _pollIntervalMs - elapsed), ct).ConfigureAwait(false);
                }
            }
            catch (OperationCanceledException)
            {
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

            byte[]? buf = SendData(new byte[0], size, rts: true);

            if (buf == null) return false;
            return System.Text.Encoding.ASCII.GetString(buf, 0, size).TrimEnd('\0').CompareTo(WelcomeMessage) == 0;
        }

        private VendorDataStruct? ReadVendorData()
        {
            if (_port == null) return null;

            var size = Marshal.SizeOf<VendorDataStruct>();
            byte[]? buf = SendCmd(UsbCmd.CMD_READ_VENDOR_DATA, size);

            if (buf == null) return null;
            return BytesToStruct<VendorDataStruct>(buf);
        }

        private byte[]? ReadUidBytes()
        {
            if (_port == null) return null;

            const int uidBytes = 12;
            return SendCmd(UsbCmd.CMD_READ_UID, uidBytes);
        }

        // The Phanteks firmware has been seen with both the compact (96 B) and the fan-field (104 B)
        // sensor layout, so measure the actual reply size once instead of trusting HasFanSensors.
        private void DetectSensorLayout()
        {
            if (_port == null) return;

            _hasFanSensorsDetected = null;

            lock (_port)
            {
                _port!.Open();
                try
                {
                    _port!.DiscardInBuffer();
                    _port!.Write(new[] { (byte)UsbCmd.CMD_READ_SENSOR_VALUES }, 0, 1);

                    var buf = new byte[SensorStructSize + 16];
                    int total = 0;
                    int previousTimeout = _port!.ReadTimeout;
                    _port!.ReadTimeout = 250;
                    try
                    {
                        while (total < buf.Length)
                        {
                            int read;
                            try
                            {
                                read = _port!.Read(buf, total, buf.Length - total);
                            }
                            catch (TimeoutException)
                            {
                                break;
                            }

                            if (read <= 0) break;
                            total += read;
                        }
                    }
                    finally
                    {
                        _port!.ReadTimeout = previousTimeout;
                    }

                    if (total == SensorStructSize) _hasFanSensorsDetected = true;
                    else if (total == SensorStructNoFanSize) _hasFanSensorsDetected = false;
                }
                finally
                {
                    _port!.Close();
                }
            }
        }

        private SensorStruct? ReadSensorValues()
        {
            if (_port == null) return null;

            if (!UseFanSensorLayout)
            {
                byte[]? compactBuf = SendCmd(UsbCmd.CMD_READ_SENSOR_VALUES, SensorStructNoFanSize);
                if (compactBuf == null) return null;

                var s = BytesToStruct<SensorStructNoFan>(compactBuf);
                return new SensorStruct
                {
                    Ts = new short[] { s.Ts[0], s.Ts[1], 0, 0 },
                    Vdd = s.Vdd,
                    PowerReadings = s.PowerReadings,
                    TotalPower = s.TotalPower,
                    TotalCurrent = s.TotalCurrent,
                    AvgVoltage = s.AvgVoltage,
                    HpwrCapability = s.HpwrCapability,
                    FaultStatus = s.FaultStatus,
                    FaultLog = s.FaultLog
                };
            }

            byte[]? buf = SendCmd(UsbCmd.CMD_READ_SENSOR_VALUES, SensorStructSize);

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
                OnboardTempInC = ss.Ts[(int)SensorTs.SENSOR_TS1] / 10.0,
                OnboardTempOutC = ss.Ts[(int)SensorTs.SENSOR_TS2] / 10.0,
                ExternalTemp1C = -100,
                ExternalTemp2C = -100,
                PsuCapabilityW = ss.HpwrCapability == HpwrCapability.PSU_CAP_600W ? 600 :
                                  ss.HpwrCapability == HpwrCapability.PSU_CAP_450W ? 450 :
                                  ss.HpwrCapability == HpwrCapability.PSU_CAP_300W ? 300 :
                                  ss.HpwrCapability == HpwrCapability.PSU_CAP_150W ? 150 : 0,

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

        private int? ReadConfigVersion()
        {
            if (_port == null) return null;
            byte[]? buf = SendCmd(UsbCmd.CMD_READ_CONFIG, 4);
            if (buf == null) return null;
            return buf[2];
        }

        private byte[]? SendCmd(UsbCmd cmd, int responseSize = 0, bool rts = false)
        {
            return SendData(new[] { (byte)cmd }, responseSize, rts);
        }

        private byte[]? SendData(byte[] data, int responseSize = 0, bool rts = false)
        {
            if (_port == null) return null;
            byte[]? buf = null;
            lock (_port)
            {
                _port!.Open();
                try
                {
                    _port!.DiscardInBuffer();
                    if (rts)
                    {
                        _port!.RtsEnable = true;
                    }
                    _port!.Write(data, 0, data.Length);
                    if (responseSize > 0)
                    {
                        buf = ReadExact(responseSize);
                    }
                    if (rts)
                    {
                        _port!.RtsEnable = false;
                    }
                }
                finally
                {
                    _port!.Close();
                }
            }
            return buf;
        }

        private byte[]? ReadExact(int size)
        {
            var buf = new byte[size];
            int offset = 0;
            int timeout = 1000;
            var start = Environment.TickCount64;

            while (offset < size && Environment.TickCount64 - start < timeout)
            {
                int read;
                try
                {
                    // Blocking read; spinning on BytesToRead would pin a core for the whole transfer.
                    read = _port!.Read(buf, offset, size - offset);
                }
                catch (TimeoutException)
                {
                    break;
                }

                if (read <= 0) break;
                offset += read;
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