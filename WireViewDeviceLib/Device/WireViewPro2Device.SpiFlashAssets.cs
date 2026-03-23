namespace WireView2.Device;

public partial class WireViewPro2Device
{
    // ── Enums ────────────────────────────────────────────────────────────────

    public enum THEME_BACKGROUND : byte
    {
        ThermalGrizzlyOrange = 1,
        ThermalGrizzlyDark   = 2,
        Disabled             = 255
    }

    public enum THEME_FAN : byte
    {
        ThermalGrizzlyOrange     = 0x64, // Bitmap 4 + 6
        ThermalGrizzlyDark       = 0x75, // Bitmap 5 + 7
        ThermalGrizzlyBlackWhite = 0x98, // Bitmap 8 + 9
    }

    // ── Theme Color Constants ─────────────────────────────────────────────────

    private const UInt32 THEME_PRIMARY_COLOR_TG1    = 0xFFFFFFFF; // White
    private const UInt32 THEME_SECONDARY_COLOR_TG1  = 0xFF646464; // Light gray
    private const UInt32 THEME_HIGHLIGHT_COLOR_TG1  = 0xFFE64121; // TG orange
    private const UInt32 THEME_BACKGROUND_COLOR_TG1 = 0xFF000000; // Black

    private const UInt32 THEME_PRIMARY_COLOR_TG2    = 0xFFFFFFFF; // White
    private const UInt32 THEME_SECONDARY_COLOR_TG2  = 0xFF646464; // Light gray
    private const UInt32 THEME_HIGHLIGHT_COLOR_TG2  = 0xFFBEBEBE; // TG light gray
    private const UInt32 THEME_BACKGROUND_COLOR_TG2 = 0xFF000000; // Black

    private const UInt32 THEME_PRIMARY_COLOR_TG3    = 0xFF969696; // Grey
    private const UInt32 THEME_SECONDARY_COLOR_TG3  = 0xFF505050; // Lighter gray
    private const UInt32 THEME_HIGHLIGHT_COLOR_TG3  = 0xFFFFFFFF; // White
    private const UInt32 THEME_BACKGROUND_COLOR_TG3 = 0xFF000000; // Black

    // ── Flash Layout Constants ────────────────────────────────────────────────

    private const UInt32 SpiFlashAssetStartOffset = 3u * 4u * 1024u; // DATA_READER_ADDR_OFFSET (3 sectors)

    private const UInt32 THEME_BACKGROUND_OFFSET_TG1 = 0x00000000;
    private const UInt32 THEME_BACKGROUND_OFFSET_TG2 = 0x0001A900;
    private const UInt32 THEME_BACKGROUND_SIZE        = 0x1A900;

    private const UInt32 THEME_FAN1_OFFSET_TG1 = 0x00053374;
    private const UInt32 THEME_FAN2_OFFSET_TG1 = 0x000586BC;
    private const UInt32 THEME_FAN1_OFFSET_TG2 = 0x00055D18;
    private const UInt32 THEME_FAN2_OFFSET_TG2 = 0x0005B060;
    private const UInt32 THEME_FAN1_OFFSET_TG3 = 0x0005DA04;
    private const UInt32 THEME_FAN2_OFFSET_TG3 = 0x000603A8;
    private const UInt32 THEME_FAN_SIZE         = 0x29A2;

    public const int ThemeBackgroundWidth  = 320;
    public const int ThemeBackgroundHeight = 170;
    public const int ThemeFanWidth  = 73;
    public const int ThemeFanHeight = 73;

    // ── Low-level Asset Read ──────────────────────────────────────────────────

    public async Task<byte[]> ReadAssetFromSpiFlashAsync(uint assetStartAddr, uint assetSize, CancellationToken ct)
    {
        var result = new byte[assetSize];
        uint bytesRead = 0;

        while (bytesRead < assetSize)
        {
            uint currentAddr    = assetStartAddr + bytesRead;
            uint remainingBytes = assetSize - bytesRead;
            uint readLen        = Math.Min(remainingBytes, SpiFlashMaxReadLen);

            byte[] pageData = await SpiFlashReadPageAsync(currentAddr, readLen, ct).ConfigureAwait(false);

            // FIX #1: Guard against infinite loop if SpiFlashReadPageAsync returns 0 bytes
            if (pageData.Length == 0)
                throw new InvalidOperationException(
                    $"SpiFlashReadPageAsync returned 0 bytes at address 0x{currentAddr:X8}. " +
                    "Possible hardware or driver issue.");

            Array.Copy(pageData, 0, result, bytesRead, pageData.Length);
            bytesRead += (uint)pageData.Length;
        }

        return result;
    }

    // ── Offset Resolver ───────────────────────────────────────────────────────

    private static uint GetThemeBackgroundOffset(THEME_BACKGROUND background)
        => background switch
        {
            THEME_BACKGROUND.ThermalGrizzlyOrange => THEME_BACKGROUND_OFFSET_TG1 + SpiFlashAssetStartOffset,
            THEME_BACKGROUND.ThermalGrizzlyDark   => THEME_BACKGROUND_OFFSET_TG2 + SpiFlashAssetStartOffset,
            _ => throw new ArgumentOutOfRangeException(nameof(background), background, "Unsupported theme background")
        };

    private static (uint Frame1Offset, uint Frame2Offset) GetThemeFanOffsets(THEME_FAN fan)
        => fan switch
        {
            THEME_FAN.ThermalGrizzlyOrange     => (THEME_FAN1_OFFSET_TG1 + SpiFlashAssetStartOffset, THEME_FAN2_OFFSET_TG1 + SpiFlashAssetStartOffset),
            THEME_FAN.ThermalGrizzlyDark       => (THEME_FAN1_OFFSET_TG2 + SpiFlashAssetStartOffset, THEME_FAN2_OFFSET_TG2 + SpiFlashAssetStartOffset),
            THEME_FAN.ThermalGrizzlyBlackWhite => (THEME_FAN1_OFFSET_TG3 + SpiFlashAssetStartOffset, THEME_FAN2_OFFSET_TG3 + SpiFlashAssetStartOffset),
            _ => throw new ArgumentOutOfRangeException(nameof(fan), fan, "Unsupported theme fan")
        };

    // ── Public Read API ───────────────────────────────────────────────────────

    public Task<byte[]?> ReadThemeBackgroundRgb565Async(THEME_BACKGROUND background, CancellationToken ct = default)
    {
        if (background == THEME_BACKGROUND.Disabled)
            return Task.FromResult<byte[]?>(null);

        var offset = GetThemeBackgroundOffset(background);
        return ReadSpiFlashBytesAsync(offset, THEME_BACKGROUND_SIZE, progress: null, ct)!;
    }

    public async Task<(byte[] Frame1, byte[] Frame2)> ReadThemeFanRgb565Async(THEME_FAN fan, CancellationToken ct = default)
    {
        var (f1, f2) = GetThemeFanOffsets(fan);
        var frame1 = await ReadSpiFlashBytesAsync(f1, THEME_FAN_SIZE, progress: null, ct).ConfigureAwait(false);
        var frame2 = await ReadSpiFlashBytesAsync(f2, THEME_FAN_SIZE, progress: null, ct).ConfigureAwait(false);
        return (frame1, frame2);
    }

    // ── Public Write API ──────────────────────────────────────────────────────

    public Task WriteThemeBackgroundRgb565Async(
        THEME_BACKGROUND background,
        ReadOnlyMemory<byte> rgb565Bytes,
        IProgress<double>? progress = null,
        CancellationToken ct = default)
    {
        if (background == THEME_BACKGROUND.Disabled)
            throw new ArgumentOutOfRangeException(nameof(background), "Cannot write Disabled background.");

        if (rgb565Bytes.Length != THEME_BACKGROUND_SIZE)
            throw new ArgumentException(
                $"Background must be exactly {THEME_BACKGROUND_SIZE} bytes (RGB565 {ThemeBackgroundWidth}x{ThemeBackgroundHeight}).",
                nameof(rgb565Bytes));

        var offset = GetThemeBackgroundOffset(background);
        return WriteSpiFlashBytesPreserveSectorsAsync(offset, rgb565Bytes, progress, ct);
    }

    public Task WriteThemeFanRgb565Async(
        THEME_FAN fan,
        ReadOnlyMemory<byte> frame1Rgb565Bytes,
        ReadOnlyMemory<byte> frame2Rgb565Bytes,
        IProgress<double>? progress = null,
        CancellationToken ct = default)
    {
        if (frame1Rgb565Bytes.Length != THEME_FAN_SIZE)
            throw new ArgumentException(
                $"Fan frame must be exactly {THEME_FAN_SIZE} bytes (RGB565 {ThemeFanWidth}x{ThemeFanHeight}).",
                nameof(frame1Rgb565Bytes));

        if (frame2Rgb565Bytes.Length != THEME_FAN_SIZE)
            throw new ArgumentException(
                $"Fan frame must be exactly {THEME_FAN_SIZE} bytes (RGB565 {ThemeFanWidth}x{ThemeFanHeight}).",
                nameof(frame2Rgb565Bytes));

        var (f1, f2) = GetThemeFanOffsets(fan);
        return WriteThemeFanInternalAsync(f1, f2, frame1Rgb565Bytes, frame2Rgb565Bytes, progress, ct);
    }

    // ── Internal Write Helpers ────────────────────────────────────────────────

    private async Task WriteThemeFanInternalAsync(
        uint frame1Addr,
        uint frame2Addr,
        ReadOnlyMemory<byte> frame1Rgb565Bytes,
        ReadOnlyMemory<byte> frame2Rgb565Bytes,
        IProgress<double>? progress,
        CancellationToken ct)
    {
        // FIX #4: Clamp p1 to max 0.5 so overall progress never jumps backwards when p2 starts
        var p1 = progress is null ? null : new Progress<double>(p => progress.Report(Math.Min(p * 0.5, 0.5)));
        var p2 = progress is null ? null : new Progress<double>(p => progress.Report(0.5 + (p * 0.5)));

        await WriteSpiFlashBytesPreserveSectorsAsync(frame1Addr, frame1Rgb565Bytes, p1, ct).ConfigureAwait(false);
        await WriteSpiFlashBytesPreserveSectorsAsync(frame2Addr, frame2Rgb565Bytes, p2, ct).ConfigureAwait(false);
    }
}
