using System;
using System.Threading;
using System.Threading.Tasks;

namespace WireView2.Device
{
    public sealed class DeviceAutoConnector : IDisposable
    {
        // Shared singleton for the whole app
        public static DeviceAutoConnector Shared { get; } = new DeviceAutoConnector();

        private readonly object _gate = new();
        private CancellationTokenSource? _cts;
        private Task? _worker;

        private IWireViewDevice? _device;
        private int _pollMs = 1000;

        public event EventHandler<bool>? ConnectionChanged; // true=connected
        public event EventHandler<DeviceData>? DataUpdated;

        // Keep the handler we attach so we can detach the exact same delegate
        private EventHandler<DeviceData>? _dataForwardHandler;

        public IWireViewDevice? Device
        {
            get
            {
                lock (_gate)
                {
                    return _device;
                }
            }
        }

        public void Start()
        {
            if (_worker != null) return;
            _cts = new CancellationTokenSource();
            _worker = Task.Run(() => LoopAsync(_cts.Token));
        }

        public void Stop()
        {
            _cts?.Cancel();
            try { _worker?.Wait(500); } catch { }
            _worker = null;
            _cts = null;
            DisconnectInternal();
        }

        public void SetPollInterval(int ms)
        {
            _pollMs = Math.Clamp(ms, 50, 5000);
            lock (_gate)
            {
                if (_device is WireViewPro2Device pro2Device) pro2Device.PollIntervalMs = _pollMs;
            }
        }

        private async Task LoopAsync(CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                try
                {
                    EnsureDevice();
                }
                catch
                {
                    // ignore and retry
                }

                await Task.Delay(_pollMs, ct).ConfigureAwait(false);
            }
        }

        private void EnsureDevice()
        {
            lock (_gate)
            {
                if (_device is { Connected: true })
                {
                    return;
                }

                // If we have a stale/disconnected instance, drop it so we can reconnect cleanly.
                if (_device != null)
                {
                    DisconnectInternal();
                }

                var ports = Stm32PortFinder.FindMatchingComPorts();
                if (ports.Count == 0)
                {
                    return;
                }

                // Try connect to all matching ports
                foreach (var port in ports)
                {
                    var basicDevice = new WireViewBasicDevice(port);
                    try
                    {
                        basicDevice.Connect();
                        if (basicDevice.Connected)
                        {
                            basicDevice.Disconnect();

                            IWireViewDevice? candidateDevice = null;

                            // Check if device matches any subtype
                            if(basicDevice.VendorId == 0xEF)
                            {
                                switch(basicDevice.ProductId)
                                {
                                    case 0x05:
                                        // WireView Pro II
                                        candidateDevice = new WireViewPro2Device(port)
                                        {
                                            PollIntervalMs = _pollMs
                                        };
                                        break;
                                    case 0x06:
                                        // WireView Pro II Noctua Edition
                                        candidateDevice = new WireViewPro2NoctuaDevice(port)
                                        {
                                            PollIntervalMs = _pollMs
                                        };
                                        break;
                                    case 0x07:
                                        // WireView II
                                        candidateDevice = new WireView2Device(port)
                                        {
                                            PollIntervalMs = _pollMs
                                        };
                                        break;
                                    case 0x08:
                                        // WireView II Phanteks Edition
                                        candidateDevice = new WireView2PhanteksDevice(port)
                                        {
                                            PollIntervalMs = _pollMs
                                        };
                                        break;
                                }
                            }
                            
                            basicDevice.Dispose();

                            if (candidateDevice == null)
                            {
                                continue;
                            }

                            _device = candidateDevice;
                            _device.Connect();

                            if (_device.Connected)
                            {
                                _device.ConnectionChanged += OnDeviceConnectionChanged;
                                _dataForwardHandler ??= (_, d) => DataUpdated?.Invoke(this, d);
                                _device.DataUpdated += _dataForwardHandler;
                                ConnectionChanged?.Invoke(this, true);
                                return;
                            }

                            _device.Disconnect();
                            if (_device is IDisposable disposableDevice) {
                                disposableDevice.Dispose();
                            }

                            _device = null;
                            continue;
                        }
                        else
                        {
                            basicDevice.Dispose();
                        }
                    }
                    catch
                    {
                        try
                        {
                            basicDevice.Dispose();
                        }
                        catch
                        {
                        }
                    }

                }
            }
        }

        private void OnDeviceConnectionChanged(object? sender, bool connected)
        {
            if (!connected)
            {
                // drop and let loop reconnect
                DisconnectInternal();
            }
            ConnectionChanged?.Invoke(this, connected);
        }

        private void DisconnectInternal()
        {
            try
            {
                if (_device != null)
                {
                    _device.ConnectionChanged -= OnDeviceConnectionChanged;

                    if (_dataForwardHandler != null)
                        _device.DataUpdated -= _dataForwardHandler;

                    _device.Disconnect();
                    if(_device is WireViewPro2Device pro2Device)
                    {
                        pro2Device.Dispose();
                    } else if(_device is WireViewPro2NoctuaDevice noctuaPro2Device)
                    {
                        noctuaPro2Device.Dispose();
                    } else if (_device is WireViewBasicDevice basicDevice)
                    {
                        basicDevice.Dispose();
                    }
                    _device = null;
                }
            }
            catch { }
            finally
            {
                _device = null;
                ConnectionChanged?.Invoke(this, false);
            }
        }

        public void Dispose() => Stop();
    }
}