using System.Diagnostics;
using System.Text.RegularExpressions;

namespace WireView2.Device
{
    /// <summary>
    /// Windows-only helper for SetupAPI-based device presence checks and PnPUtil-based driver store operations.
    /// </summary>
    public static class WindowsDriverHelper
    {
        // WinUSB device interface class GUID (GUID_DEVINTERFACE_WINUSB)
        private static readonly Guid GuidDevInterfaceWinUsb = new("dee824ef-729b-4a0e-9c14-b7117d33a817");

        public static async Task<int> StopServicesMatchingTmInstallAsync(CancellationToken cancellationToken = default)
        {
            if (!OperatingSystem.IsWindows())
            {
                return 0;
            }

            // Query all services, then stop those whose SERVICE_NAME matches: tm*Install
            // Example matches: tmFooInstall, tmInstall, tm123Install
            var queryPsi = new ProcessStartInfo
            {
                FileName = "sc.exe",
                Arguments = "query state= all",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var queryProc = Process.Start(queryPsi) ?? throw new InvalidOperationException("Failed to start sc.exe.");
            var outputTask = queryProc.StandardOutput.ReadToEndAsync(cancellationToken);
            var errorTask = queryProc.StandardError.ReadToEndAsync(cancellationToken);

            await queryProc.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
            var output = await outputTask.ConfigureAwait(false);
            _ = await errorTask.ConfigureAwait(false);

            if (queryProc.ExitCode != 0)
            {
                return 0;
            }

            var servicesToStop = new List<string>();
            var rx = new Regex(@"^\s*SERVICE_NAME\s*:\s*(?<name>\S+)\s*$", RegexOptions.Multiline | RegexOptions.CultureInvariant);
            foreach (Match m in rx.Matches(output))
            {
                cancellationToken.ThrowIfCancellationRequested();

                var name = m.Groups["name"].Value;
                if (name.StartsWith("tm", StringComparison.OrdinalIgnoreCase) &&
                    name.EndsWith("Install", StringComparison.OrdinalIgnoreCase))
                {
                    servicesToStop.Add(name);
                }
            }

            var unique = servicesToStop.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            int stoppedCount = 0;

            foreach (var serviceName in unique)
            {
                cancellationToken.ThrowIfCancellationRequested();

                // Stop may require admin privileges depending on service permissions.
                var stopPsi = new ProcessStartInfo
                {
                    FileName = "sc.exe",
                    Arguments = $"stop \"{serviceName}\"",
                    UseShellExecute = true,
                    Verb = "runas"
                };

                using var stopProc = Process.Start(stopPsi);
                if (stopProc is null)
                {
                    continue;
                }

                await stopProc.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
                if (stopProc.ExitCode == 0)
                {
                    stoppedCount++;
                }
            }

            return stoppedCount;
        }

        public static async Task<bool> WaitForDevicePresentAsync(ushort vid, ushort pid, TimeSpan timeout, CancellationToken cancellationToken = default)
        {
            try
            {
                var end = DateTime.UtcNow + timeout;
                while (DateTime.UtcNow < end)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    if (IsDevicePresent(vid, pid))
                    {
                        return true;
                    }

                    await Task.Delay(1000, cancellationToken).ConfigureAwait(false);
                }

                return false;
            }
            catch
            {
                // Treat any SetupAPI/Interop failure as "not available" rather than crashing the app.
                return false;
            }
        }

        public static async Task<bool> WaitForWinUsbDeviceInterfaceAsync(ushort vid, ushort pid, TimeSpan timeout, CancellationToken cancellationToken = default)
        {
            try
            {
                var end = DateTime.UtcNow + timeout;
                while (DateTime.UtcNow < end)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    if (IsWinUsbDeviceInterfacePresent(vid, pid))
                    {
                        return true;
                    }

                    await Task.Delay(1000, cancellationToken).ConfigureAwait(false);
                }

                return false;
            }
            catch
            {
                // Treat any SetupAPI/Interop failure as "not available" rather than crashing the app.
                return false;
            }
        }

        public static bool IsDevicePresent(ushort vid, ushort pid)
        {
            if (!OperatingSystem.IsWindows())
            {
                return false;
            }

            var needle = $"vid_{vid:X4}&pid_{pid:X4}";
            return WindowsSetupApi.EnumerateDeviceInstanceIds().Any(id =>
                id.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        public static bool IsWinUsbDeviceInterfacePresent(ushort vid, ushort pid)
        {
            if (!OperatingSystem.IsWindows())
            {
                return false;
            }

            var needle = $"vid_{vid:X4}&pid_{pid:X4}";
            foreach (var devicePath in WindowsSetupApi.EnumerateDeviceInterfacePaths(GuidDevInterfaceWinUsb))
            {
                if (devicePath.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        public static string? TryGetDeviceDescription(ushort vid, ushort pid)
        {
            if (!OperatingSystem.IsWindows())
            {
                return null;
            }

            var needle = $"vid_{vid:X4}&pid_{pid:X4}";

            try
            {
                foreach (var (instanceId, deviceDesc, friendlyName) in WindowsSetupApi.EnumeratePresentDevicesWithNames())
                {
                    if (instanceId.IndexOf(needle, StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        continue;
                    }

                    // Prefer FriendlyName when present; fall back to DeviceDesc.
                    return deviceDesc;
                }
            }
            catch
            {
                // Match presence-check behavior: treat interop failure as "unknown".
            }

            return null;
        }

        public static async Task<bool> EnsureDriverInstalledAsync(
            string infPath,
            CancellationToken cancellationToken = default)
        {
            if (!OperatingSystem.IsWindows())
            {
                return false;
            }

            if (!File.Exists(infPath))
            {
                throw new FileNotFoundException("Driver INF not found.", infPath);
            }

            // Requires admin. If not elevated, pnputil will fail (access denied).
            var psi = new ProcessStartInfo
            {
                FileName = "pnputil.exe",
                Arguments = $"/add-driver \"{infPath}\" /install",
                UseShellExecute = true,
                Verb = "runas",
            };

            using var proc = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start pnputil.exe.");

            while (!proc.HasExited)
            {
                cancellationToken.ThrowIfCancellationRequested();
                await Task.Delay(100, cancellationToken).ConfigureAwait(false);
            }

            return proc.ExitCode == 0;
        }

        public static async Task<bool> IsDriverInfInstalledAsync(
            string infPath,
            CancellationToken cancellationToken = default)
        {
            if (!OperatingSystem.IsWindows())
            {
                return false;
            }

            if (!File.Exists(infPath))
            {
                throw new FileNotFoundException("Driver INF not found.", infPath);
            }

            // Does not require admin
            var psi = new ProcessStartInfo
            {
                FileName = "pnputil.exe",
                Arguments = "/enum-drivers",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };

            using var proc = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start pnputil.exe.");
            string output = await proc.StandardOutput.ReadToEndAsync().ConfigureAwait(false);

            while (!proc.HasExited)
            {
                cancellationToken.ThrowIfCancellationRequested();
                await Task.Delay(50, cancellationToken).ConfigureAwait(false);
            }

            if (proc.ExitCode != 0)
            {
                throw new InvalidOperationException($"Driver enumeration failed with exit code {proc.ExitCode}.");
            }

            string needle = Path.GetFileNameWithoutExtension(infPath);
            return output.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static Task<bool> IsDriverInstalledByOriginalInfNameAsync(string originalInfFileName, CancellationToken cancellationToken = default) =>
            IsPnpDriverInstalledAsync(originalInfFileName, cancellationToken);

        public static async Task<bool> RemoveDriverByOriginalInfNameIfPresentAsync(string originalInfFileName, CancellationToken cancellationToken = default)
        {
            if (!OperatingSystem.IsWindows())
            {
                return false;
            }

            var publishedNames = await FindPublishedOemInfNamesByOriginalInfAsync(originalInfFileName, cancellationToken).ConfigureAwait(false);
            if (publishedNames.Count == 0)
            {
                return false;
            }

            bool removedAny = false;
            foreach (var publishedName in publishedNames)
            {
                if (await DeleteDriverByPublishedNameAsync(publishedName, cancellationToken).ConfigureAwait(false))
                {
                    removedAny = true;
                }
            }

            return removedAny;
        }

        private static async Task<bool> IsPnpDriverInstalledAsync(string infFileName, CancellationToken cancellationToken)
        {
            if (!OperatingSystem.IsWindows())
            {
                return false;
            }

            var psi = new ProcessStartInfo
            {
                FileName = "pnputil.exe",
                Arguments = "/enum-drivers",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };

            using var proc = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start pnputil.exe.");
            string output = await proc.StandardOutput.ReadToEndAsync().ConfigureAwait(false);

            while (!proc.HasExited)
            {
                cancellationToken.ThrowIfCancellationRequested();
                await Task.Delay(50, cancellationToken).ConfigureAwait(false);
            }

            if (proc.ExitCode != 0)
            {
                throw new InvalidOperationException($"Driver enumeration failed with exit code {proc.ExitCode}.");
            }

            return output.IndexOf(infFileName, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static async Task<List<string>> FindPublishedOemInfNamesByOriginalInfAsync(string originalInfFileName, CancellationToken cancellationToken)
        {
            var psi = new ProcessStartInfo
            {
                FileName = "pnputil.exe",
                Arguments = "/enum-drivers",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };

            using var proc = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start pnputil.exe.");
            string output = await proc.StandardOutput.ReadToEndAsync().ConfigureAwait(false);

            while (!proc.HasExited)
            {
                cancellationToken.ThrowIfCancellationRequested();
                await Task.Delay(50, cancellationToken).ConfigureAwait(false);
            }

            if (proc.ExitCode != 0)
            {
                throw new InvalidOperationException($"Driver enumeration failed with exit code {proc.ExitCode}.");
            }

            // pnputil output includes blocks like:
            //   Published Name : oemXX.inf
            //   Original Name  : guistdfudev.inf
            var lines = output.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

            string? published = null;
            var results = new List<string>();

            foreach (var rawLine in lines)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var line = rawLine.Trim();

                const string publishedPrefix = "Published Name";
                const string originalPrefix = "Original Name";

                if (line.StartsWith(publishedPrefix, StringComparison.OrdinalIgnoreCase))
                {
                    published = ExtractPnputilValue(line);
                    continue;
                }

                if (line.StartsWith(originalPrefix, StringComparison.OrdinalIgnoreCase))
                {
                    var original = ExtractPnputilValue(line);
                    if (published is not null &&
                        original is not null &&
                        original.Equals(originalInfFileName, StringComparison.OrdinalIgnoreCase))
                    {
                        results.Add(published);
                    }

                    published = null;
                }
            }

            return results;
        }

        private static string? ExtractPnputilValue(string line)
        {
            int idx = line.IndexOf(':');
            if (idx < 0 || idx == line.Length - 1)
            {
                return null;
            }

            return line[(idx + 1)..].Trim();
        }

        private static async Task<bool> DeleteDriverByPublishedNameAsync(string publishedInfName, CancellationToken cancellationToken)
        {
            // Typically requires admin; if not elevated, pnputil will fail (access denied).
            var psi = new ProcessStartInfo
            {
                FileName = "pnputil.exe",
                Arguments = $"/delete-driver \"{publishedInfName}\" /uninstall /force",
                UseShellExecute = true,
                Verb = "runas",
            };

            using var proc = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start pnputil.exe.");

            while (!proc.HasExited)
            {
                cancellationToken.ThrowIfCancellationRequested();
                await Task.Delay(100, cancellationToken).ConfigureAwait(false);
            }

            return proc.ExitCode == 0;
        }
    }
}