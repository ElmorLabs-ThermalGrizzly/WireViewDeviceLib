using System.Runtime.InteropServices;

namespace WireView2.Device
{
    public static class Stm32PortFinder
    {
        // SetupAPI constants (subset)
        private const uint DIGCF_PRESENT = 0x00000002;
        private const uint DIGCF_ALLCLASSES = 0x00000004;

        private const uint SPDRP_HARDWAREID = 0x00000001;
        private const uint SPDRP_FRIENDLYNAME = 0x0000000C;

        private static readonly nint INVALID_HANDLE_VALUE = new(-1);

        public static List<string> FindMatchingComPorts()
        {
            var ports = new List<string>();

            if (!OperatingSystem.IsWindows())
            {
                return ports;
            }

            nint devInfo = SetupDiGetClassDevsW(IntPtr.Zero, null, IntPtr.Zero, DIGCF_PRESENT | DIGCF_ALLCLASSES);
            if (devInfo == INVALID_HANDLE_VALUE)
            {
                return ports;
            }

            try
            {
                var devInfoData = new SP_DEVINFO_DATA
                {
                    cbSize = (uint)Marshal.SizeOf<SP_DEVINFO_DATA>()
                };

                for (uint index = 0; SetupDiEnumDeviceInfo(devInfo, index, ref devInfoData); index++)
                {
                    var hardwareIds = TryGetDeviceRegistryPropertyMultiSz(devInfo, ref devInfoData, SPDRP_HARDWAREID);
                    if (hardwareIds is null || hardwareIds.Length == 0)
                    {
                        continue;
                    }

                    if (!HardwareIdsContainVidPid(hardwareIds, "VID_0483", "PID_5740"))
                    {
                        continue;
                    }

                    var friendlyName = TryGetDeviceRegistryPropertyString(devInfo, ref devInfoData, SPDRP_FRIENDLYNAME);
                    if (string.IsNullOrWhiteSpace(friendlyName))
                    {
                        continue;
                    }

                    var comPort = TryExtractComPortFromFriendlyName(friendlyName);
                    if (!string.IsNullOrWhiteSpace(comPort))
                    {
                        ports.Add(comPort);
                    }
                }
            }
            finally
            {
                _ = SetupDiDestroyDeviceInfoList(devInfo);
            }

            return ports;
        }

        private static bool HardwareIdsContainVidPid(string[] hardwareIds, string vid, string pid)
        {
            for (int i = 0; i < hardwareIds.Length; i++)
            {
                var s = hardwareIds[i];
                if (string.IsNullOrWhiteSpace(s))
                {
                    continue;
                }

                if (s.Contains(vid, StringComparison.OrdinalIgnoreCase) &&
                    s.Contains(pid, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static string? TryExtractComPortFromFriendlyName(string friendlyName)
        {
            var start = friendlyName.LastIndexOf("(COM", StringComparison.OrdinalIgnoreCase);
            if (start < 0)
            {
                return null;
            }

            var end = friendlyName.IndexOf(')', start);
            if (end <= start)
            {
                return null;
            }

            return friendlyName.Substring(start + 1, end - start - 1);
        }

        private static string? TryGetDeviceRegistryPropertyString(nint devInfoSet, ref SP_DEVINFO_DATA devInfoData, uint property)
        {
            _ = SetupDiGetDeviceRegistryPropertyW(
                devInfoSet,
                ref devInfoData,
                property,
                out _,
                null,
                0,
                out uint requiredSize);

            if (requiredSize == 0)
            {
                return null;
            }

            var buffer = new byte[requiredSize];
            if (!SetupDiGetDeviceRegistryPropertyW(
                devInfoSet,
                ref devInfoData,
                property,
                out _,
                buffer,
                (uint)buffer.Length,
                out _))
            {
                return null;
            }

            var s = EncodingFromUtf16Bytes(buffer);
            return string.IsNullOrWhiteSpace(s) ? null : s;
        }

        private static string[]? TryGetDeviceRegistryPropertyMultiSz(nint devInfoSet, ref SP_DEVINFO_DATA devInfoData, uint property)
        {
            _ = SetupDiGetDeviceRegistryPropertyW(
                devInfoSet,
                ref devInfoData,
                property,
                out _,
                null,
                0,
                out uint requiredSize);

            if (requiredSize == 0)
            {
                return null;
            }

            var buffer = new byte[requiredSize];
            if (!SetupDiGetDeviceRegistryPropertyW(
                devInfoSet,
                ref devInfoData,
                property,
                out _,
                buffer,
                (uint)buffer.Length,
                out _))
            {
                return null;
            }

            return DecodeMultiSzUtf16(buffer);
        }

        private static string EncodingFromUtf16Bytes(byte[] bytes)
        {
            var s = System.Text.Encoding.Unicode.GetString(bytes);
            var nul = s.IndexOf('\0', StringComparison.Ordinal);
            return nul >= 0 ? s.Substring(0, nul) : s;
        }

        private static string[] DecodeMultiSzUtf16(byte[] bytes)
        {
            var s = System.Text.Encoding.Unicode.GetString(bytes);
            return s.Split('\0', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SP_DEVINFO_DATA
        {
            public uint cbSize;
            public Guid ClassGuid;
            public uint DevInst;
            public nint Reserved;
        }

        [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern nint SetupDiGetClassDevsW(
            nint classGuid,
            string? enumerator,
            nint hwndParent,
            uint flags);

        [DllImport("setupapi.dll", SetLastError = true)]
        private static extern bool SetupDiEnumDeviceInfo(
            nint deviceInfoSet,
            uint memberIndex,
            ref SP_DEVINFO_DATA deviceInfoData);

        [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SetupDiGetDeviceRegistryPropertyW(
            nint deviceInfoSet,
            ref SP_DEVINFO_DATA deviceInfoData,
            uint property,
            out uint propertyRegDataType,
            byte[]? propertyBuffer,
            uint propertyBufferSize,
            out uint requiredSize);

        [DllImport("setupapi.dll", SetLastError = true)]
        private static extern bool SetupDiDestroyDeviceInfoList(nint deviceInfoSet);
    }
}