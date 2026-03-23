using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace WireView2.Device
{
    /// <summary>
    /// Windows-only helper for SetupAPI-based device presence checks and PnPUtil-based driver store operations.
    /// </summary>
    public static class WindowsDriverHelper
    {
        // WinUSB device interface class GUID (GUID_DEVINTERFACE_WINUSB)
        private static readonly Guid GuidDevInterfaceWinUsb = new("dee824ef-729b-4a0e-9c14-b7117d33a817");

        public static async Task<IReadOnlyList<string>> StopServicesMatchingTmInstallAsync(CancellationToken cancellationToken = default)
        {
            if (!OperatingSystem.IsWindows())
            {
                return Array.Empty<string>();
            }

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
                return Array.Empty<string>();
            }

            var servicesToStop = new List<string>();
            
            // FIX: Supports "SERVICE_NAME" (English) and "DIENST_NAME"/"DIENSTNAME" (German)
            var rx = new Regex(@"^\s*(SERVICE_NAME|DIENST_NAME|DIENSTNAME)\s*:\s*(?<name>\S+)\s*$", RegexOptions.Multiline | RegexOptions.CultureInvariant);
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
            var stopped = new List<string>(capacity: unique.Count);

            foreach (var serviceName in unique)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var stopPsi = new ProcessStartInfo
                {
                    FileName = "sc.exe",
                    Arguments = $"stop \"{serviceName}\"",
                    UseShellExecute = true,
                    CreateNoWindow = true,
                    Verb = "runas"
                };

                try
                {
                    using var stopProc = Process.Start(stopPsi);
                    if (stopProc is null) continue;

                    await stopProc.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
                    if (stopProc.ExitCode == 0)
                    {
                        stopped.Add(serviceName);
                    }
                }
                catch (Win32Exception)
                {
                    // UAC prompt (admin rights) was declined by the user
                    continue;
                }
            }

            return stopped;
        }

        public static async Task<int> StartServicesAsync(IEnumerable<string> serviceNames, CancellationToken cancellationToken = default)
        {
            if (!OperatingSystem.IsWindows())
            {
                return 0;
            }

            int startedCount = 0;

            foreach (var serviceName in serviceNames.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                cancellationToken.ThrowIfCancellationRequested();

                var startPsi = new ProcessStartInfo
                {
                    FileName = "sc.exe",
                    Arguments = $"start \"{serviceName}\"",
                    UseShellExecute = true,
                    CreateNoWindow = true,
                    Verb = "runas"
                };

                try
                {
                    using var startProc = Process.Start(startPsi);
                    if (startProc is null) continue;

                    await startProc.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
                    if (startProc.ExitCode == 0)
                    {
                        startedCount++;
                    }
                }
                catch (Win32Exception)
                {
                    // UAC prompt was declined
                    continue;
                }
            }

            return startedCount;
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
                return false;
            }
        }

        public static bool IsDevicePresent(ushort vid, ushort pid)
        {
            if (!OperatingSystem.IsWindows()) return false;

            var needle = $"vid_{vid:X4}&pid_{pid:X4}";
            
            // Note: This assumes that WindowsSetupApi.EnumerateDeviceInstanceIds() exists in the project.
            return WindowsSetupApi.EnumerateDeviceInstanceIds().Any(id =>
                id.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        public static bool IsWinUsbDeviceInterfacePresent(ushort vid, ushort pid)
        {
            if (!OperatingSystem.IsWindows()) return false;

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
            if (!OperatingSystem.IsWindows()) return null;

            var needle = $"vid_{vid:X4}&pid_{pid:X4}";

            try
            {
                foreach (var (instanceId, deviceDesc, friendlyName) in WindowsSetupApi.EnumeratePresentDevicesWithNames())
                {
                    if (instanceId.IndexOf(needle, StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        continue;
                    }

                    return friendlyName ?? deviceDesc;
                }
            }
            catch
            {
                // Ignore
            }

            return null;
        }

        public static async Task<bool> EnsureDriverInstalledAsync(string infPath, CancellationToken cancellationToken = default)
        {
            if (!OperatingSystem.IsWindows()) return false;
            if (!File.Exists(infPath)) throw new FileNotFoundException("Driver INF not found.", infPath);

            var psi = new ProcessStartInfo
            {
                FileName = "pnputil.exe",
                Arguments = $"/add-driver \"{infPath}\" /install",
                UseShellExecute = true,
                CreateNoWindow = true,
                Verb = "runas"
            };

            try
            {
                using var proc = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start pnputil.exe.");
                await proc.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
                return proc.ExitCode == 0;
            }
            catch (Win32Exception)
            {
                return false;
            }
        }

        public static async Task<bool> IsDriverInfInstalledAsync(string infPath, CancellationToken cancellationToken = default)
        {
            if (!OperatingSystem.IsWindows()) return false;
            if (!File.Exists(infPath)) throw new FileNotFoundException("Driver INF not found.", infPath);

            var psi = new ProcessStartInfo
            {
                FileName = "pnputil.exe",
                Arguments = "/enum-drivers",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };

            using var proc = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start pnputil.exe.");
            string output = await proc.StandardOutput.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
            
            await proc.WaitForExitAsync(cancellationToken).ConfigureAwait(false);

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
            if (!OperatingSystem.IsWindows()) return false;

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
            if (!OperatingSystem.IsWindows()) return false;

            var psi = new ProcessStartInfo
            {
                FileName = "pnputil.exe",
                Arguments = "/enum-drivers",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };

            using var proc = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start pnputil.exe.");
            string output = await proc.StandardOutput.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
            
            await proc.WaitForExitAsync(cancellationToken).ConfigureAwait(false);

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
            string output = await proc.StandardOutput.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
            
            await proc.WaitForExitAsync(cancellationToken).ConfigureAwait(false);

            if (proc.ExitCode != 0)
            {
                throw new InvalidOperationException($"Driver enumeration failed with exit code {proc.ExitCode}.");
            }

            var lines = output.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            string? lastOemInf = null;
            var results = new List<string>();

            foreach (var rawLine in lines)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var line = rawLine.Trim();

                // FIX: Searches language-independently for oemXX.inf
                var oemMatch = Regex.Match(line, @"oem\d+\.inf", RegexOptions.IgnoreCase);
                if (oemMatch.Success)
                {
                    lastOemInf = oemMatch.Value;
                }

                if (line.Contains(originalInfFileName, StringComparison.OrdinalIgnoreCase) && lastOemInf != null)
                {
                    results.Add(lastOemInf);
                    lastOemInf = null; // Reset for the next match
                }
            }

            return results;
        }

        private static async Task<bool> DeleteDriverByPublishedNameAsync(string publishedInfName, CancellationToken cancellationToken)
        {
            var psi = new ProcessStartInfo
            {
                FileName = "pnputil.exe",
                Arguments = $"/delete-driver \"{publishedInfName}\" /uninstall /force",
                UseShellExecute = true,
                CreateNoWindow = true,
                Verb = "runas",
            };

            try
            {
                using var proc = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start pnputil.exe.");
                await proc.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
                return proc.ExitCode == 0;
            }
            catch (Win32Exception)
            {
                // FIX: Prevents crash if the user clicks "No" in the UAC prompt
                return false;
            }
        }
    }
}
