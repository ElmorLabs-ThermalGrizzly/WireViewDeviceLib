namespace WireView2.Device;

public partial class WireViewPro2Device
{
    private const uint DataloggerStartAddr = 8u * 1024 * 1024; // second 8MB
    private const uint DataloggerEndAddr = SpiFlashLogicalSizeBytes;

    public async Task<byte[]> ReadDeviceLogAsync(
        IProgress<double>? progress = null,
        CancellationToken ct = default)
    {

        try
        {
            uint len = DataloggerEndAddr - DataloggerStartAddr;

            var bytes = await SpiFlashReadBytesAsync(
                addr: DataloggerStartAddr,
                len: len,
                progress: progress,
                ct: ct).ConfigureAwait(false);

            // Parse off-device
            return bytes;
        }
        finally
        {
            progress?.Report(1.0);
        }
    }

    public async Task<IReadOnlyList<DeviceLogParser.DATALOGGER_Entry>> ReadHistoryFromSpiFlashAsync(
        IProgress<double>? progress = null,
        CancellationToken ct = default)
    {
        try
        {
            uint len = DataloggerEndAddr - DataloggerStartAddr;

            var bytes = await SpiFlashReadBytesAsync(
                addr: DataloggerStartAddr,
                len: len,
                progress: progress,
                ct: ct).ConfigureAwait(false);

            // Parse off-device
            return DeviceLogParser.Parse(
                bytes);
        }
        finally
        {
            progress?.Report(1.0);
        }
    }
}