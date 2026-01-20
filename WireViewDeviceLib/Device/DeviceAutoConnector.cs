using System;
using System.Linq;
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

        private WireViewPro2Device? _device;
        private int _pollMs = 200;

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
                if (_device != null) _device.PollIntervalMs = _pollMs;
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

                await Task.Delay(1000, ct).ConfigureAwait(false);
            }
        }

        private void EnsureDevice()
        {
            lock (_gate)
            {
                if (_device is { Connected: true }) return;

                // Try find port and connect
                var ports = Stm32PortFinder.FindMatchingComPorts();
                foreach (var port in ports)
                {
                    if (port == null)
                    {
                        // no port present; if we had a device, dispose it
                        DisconnectInternal();
                        return;
                    }

                    // If wrong or disposed device, reset
                    if (_device == null)
                    {
                        var dev = new WireViewPro2Device(port);
                        dev.PollIntervalMs = _pollMs;
                        dev.ConnectionChanged += OnDeviceConnectionChanged;

                        _dataForwardHandler ??= (_, d) => DataUpdated?.Invoke(this, d);
                        dev.DataUpdated += _dataForwardHandler;

                        try
                        {
                            _device = dev;
                            dev.Connect();
                        }
                        catch
                        {
                            _device = null;
                            DisconnectInternal();
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
                    _device.Dispose();
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