using System.Management;
using Timer = System.Threading.Timer;

namespace WireView2.Device
{
    public class ComPortWatcher : IDisposable
    {
        private readonly Action _onDeviceChange;
        private ManagementEventWatcher? _watcher;
        private Timer? _debounceTimer;

        public ComPortWatcher(Action onDeviceChange)
        {
            _onDeviceChange = onDeviceChange;
            StartWatching();
        }

        private void StartWatching()
        {
            string query = "SELECT * FROM Win32_DeviceChangeEvent";
            _watcher = new ManagementEventWatcher(query);
            _watcher.EventArrived += (s, e) => ScheduleCheck();
            _watcher.Start();
        }

        private void ScheduleCheck()
        {
            _debounceTimer?.Dispose();
            _debounceTimer = new Timer(_ => _onDeviceChange(),
                null, 500, Timeout.Infinite);
        }

        public void Dispose()
        {
            _watcher?.Stop();
            _watcher?.Dispose();
            _debounceTimer?.Dispose();
        }
    }
}