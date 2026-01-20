using System;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
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

            return ports;
        }
    }

    public partial class WireViewPro2Device : IWireViewDevice, IDisposable
    {
        private readonly string _portName;
        private readonly int _baud;
        private SerialPort? _port;
        private CancellationTokenSource? _cts;
        private Task? _worker;

        // Serialize all command/response access to the COM port.
        private readonly SemaphoreSlim _ioLock = new(1, 1);

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

        private T? WithPortLock<T>(Func<T?> action)
        {
            _ioLock.Wait();
            try
            {
                return action();
            }
            finally
            {
                _ioLock.Release();
            }
        }

        private void WithPortLock(Action action)
        {
            _ioLock.Wait();
            try
            {
                action();
            }
            finally
            {
                _ioLock.Release();
            }
        }

        private async Task<T> WithPortLockAsync<T>(Func<Task<T>> action, CancellationToken ct)
        {
            await _ioLock.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                return await action().ConfigureAwait(false);
            }
            finally
            {
                _ioLock.Release();
            }
        }

        public void Connect()
        {
            if (Connected) return;

            _port = new SerialPort(_portName, _baud, Parity.None, 8, StopBits.One);
            _port.ReadTimeout = 1000;
            _port.WriteTimeout = 1000;
            _port.Open();

            var vd = ReadVendorData();
            if (vd != null && vd.Value.VendorId == 0xEF && vd.Value.ProductId == 0x05)
            {
                HardwareRevision = $"{vd.Value.VendorId:X2}{vd.Value.ProductId:X2}";
                FirmwareVersion = vd.Value.FwVersion.ToString();

                UniqueId = ReadUid() ?? string.Empty;

                Connected = true;
                ConnectionChanged?.Invoke(this, true);

                _cts = new CancellationTokenSource();
                _worker = Task.Run(() => PollLoop(_cts.Token));
            }
            else
            {
                Connected = false;
                _port.Close();
            }
        }

        public void Disconnect()
        {
            if (!Connected) return;

            // Ensure no in-flight transaction is using the port while closing it.
            WithPortLock(() =>
            {
                try
                {
                    _cts?.Cancel();
                    _worker?.Wait(1000);
                    _port?.Close();
                }
                catch { }
            });

            Connected = false;

            HardwareRevision = string.Empty;
            FirmwareVersion = string.Empty;
            UniqueId = string.Empty;

            ConnectionChanged?.Invoke(this, false);
        }

        public void StartSampling() { }
        public void StopSampling() { }

        public string? ReadBuildString()
        {
            if (!Connected) return null;

            return WithPortLock(() =>
            {
                _port!.DiscardInBuffer();
                _port.Write(new byte[] { (byte)UsbCmd.CMD_READ_BUILD_INFO }, 0, 1);

                var size = Marshal.SizeOf<BuildStruct>();
                var buf = ReadExact(size);
                if (buf == null) return null;

                BuildStruct buildStruct = BytesToStruct<BuildStruct>(buf);
                return buildStruct.BuildInfo;
            });
        }

        public Task EnterBootloaderAsync()
        {
            if (!Connected) return Task.CompletedTask;

            return Task.Run(() =>
            {
                try
                {
                    WithPortLock(() =>
                    {
                        _port!.DiscardInBuffer();
                        _port.Write(new byte[] { (byte)UsbCmd.CMD_BOOTLOADER }, 0, 1);
                        Thread.Sleep(50);
                    });
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
            if (!Connected) return null;

            return WithPortLock(() =>
            {
                _port!.DiscardInBuffer();
                _port.Write(new byte[] { (byte)UsbCmd.CMD_READ_CONFIG }, 0, 1);

                var size = Marshal.SizeOf<DeviceConfigStruct>();
                var buf = ReadExact(size);
                //if (buf == null) return null;
                if(buf == null) return new DeviceConfigStruct();

                return BytesToStruct<DeviceConfigStruct>(buf);
            });
        }

        public void WriteConfig(DeviceConfigStruct config)
        {
            if (!Connected) return;

            WithPortLock(() =>
            {
                var payload = StructToBytes(config);

                var frame = new byte[64];
                frame[0] = (byte)UsbCmd.CMD_WRITE_CONFIG;

                _port!.DiscardInBuffer();

                const int maxPayloadPerFrame = 62;

                for (int offset = 0; offset < payload.Length; offset += maxPayloadPerFrame)
                {
                    int bytesToWrite = Math.Min(maxPayloadPerFrame, payload.Length - offset);

                    frame[1] = (byte)offset;
                    Buffer.BlockCopy(payload, offset, frame, 2, bytesToWrite);

                    _port.Write(frame, 0, bytesToWrite + 2);
                }
            });
        }

        public void NvmCmd(NVM_CMD cmd)
        {
            if (!Connected) return;

            WithPortLock(() =>
            {
                _port!.DiscardInBuffer();
                _port.Write(new[] { (byte)UsbCmd.CMD_NVM_CONFIG, (byte)0x55, (byte)0xAA, (byte)0x55, (byte)0xAA, (byte)cmd }, 0, 6);
            });
        }

        public void ScreenCmd(SCREEN_CMD cmd)
        {
            if (!Connected) return;

            WithPortLock(() =>
            {
                _port!.DiscardInBuffer();
                _port.Write(new[] { (byte)UsbCmd.CMD_SCREEN_CHANGE, (byte)cmd }, 0, 2);
            });
        }

        public void ClearFaults(int faultStatusMask = 0xFFFF, int faultLogMask = 0xFFFF)
        {
            if (!Connected) return;
            WithPortLock(() =>
            {
                _port!.DiscardInBuffer();
                _port.Write(new[] { (byte)UsbCmd.CMD_CLEAR_FAULTS, (byte)(faultStatusMask & 0xFF), (byte)((faultStatusMask >> 8) & 0xFF), (byte)(faultLogMask & 0xFF), (byte)((faultLogMask >> 8) & 0xFF) }, 0, 5);
            });
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

        private VendorDataStruct? ReadVendorData()
        {
            return WithPortLock(() =>
            {
                _port!.DiscardInBuffer();
                _port.Write(new byte[] { (byte)UsbCmd.CMD_READ_VENDOR_DATA }, 0, 1);

                var size = Marshal.SizeOf<VendorDataStruct>();
                var buf = ReadExact(size);
                //if (buf == null) return null;
                if (buf == null) return new VendorDataStruct();
                return BytesToStruct<VendorDataStruct>(buf);
            });
        }

        private string? ReadUid()
        {
            return WithPortLock(() =>
            {
                _port!.DiscardInBuffer();
                _port.Write(new byte[] { (byte)UsbCmd.CMD_READ_UID }, 0, 1);

                const int uidBytes = 12;
                var buf = ReadExact(uidBytes);
                if (buf == null) return null;

                return BitConverter.ToString(buf).Replace("-", string.Empty);
            });
        }

        private SensorStruct? ReadSensorValues()
        {
            return WithPortLock(() =>
            {
                _port!.DiscardInBuffer();
                _port.Write(new byte[] { (byte)UsbCmd.CMD_READ_SENSOR_VALUES }, 0, 1);

                var size = Marshal.SizeOf<SensorStruct>();
                var buf = ReadExact(size);
                //if (buf == null) return null;
                if(buf == null) return new SensorStruct();

                return BytesToStruct<SensorStruct>(buf);
            });
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
                    offset += _port.Read(buf, offset, size - offset);
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