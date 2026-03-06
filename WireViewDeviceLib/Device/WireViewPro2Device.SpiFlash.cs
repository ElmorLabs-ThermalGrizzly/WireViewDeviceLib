using System;
using System.Buffers.Binary;
using System.Threading;
using System.Threading.Tasks;

namespace WireView2.Device;

public partial class WireViewPro2Device
{
    // Matches firmware
    private const uint SpiFlashSizeBytes = 0x1000000;            // 16MB physical
    private const uint SpiFlashLogicalSizeBytes = SpiFlashSizeBytes;
    private const uint SpiFlashSectorSizeBytes = 0x001000;     // 4KB
    private const uint SpiFlashPageSize = 0x100;              // 256 bytes

    private const int SpiFlashMaxReadLen = 256; // firmware page buffer size
    private const int SpiFlashMaxWriteLen = 256; // firmware page buffer size

    private const byte CmdSpiFlashReadPage = (byte)UsbCmd.CMD_SPI_FLASH_READ_PAGE;
    private const byte CmdSpiFlashWritePage = (byte)UsbCmd.CMD_SPI_FLASH_WRITE_PAGE;
    private const byte CmdSpiFlashEraseSector = (byte)UsbCmd.CMD_SPI_FLASH_ERASE_SECTOR;

    // Chunked writes avoid occasional Windows CDC semaphore timeouts when sending >64 bytes.
    private void SpiFlashWriteAllChunked(ReadOnlySpan<byte> data, int chunkSize = 64, int interChunkDelayMs = 1)
    {
        if (_port is null)
            throw new InvalidOperationException("Device not connected.");

        int offset = 0;
        while (offset < data.Length)
        {
            int n = Math.Min(chunkSize, data.Length - offset);
            var buf = data.Slice(offset, n).ToArray();
            _port.Write(buf, 0, buf.Length);
            offset += n;

            if (interChunkDelayMs > 0)
                Thread.Sleep(interChunkDelayMs);
        }
    }

    // Firmware: CMD_SPI_FLASH_READ_PAGE expects payload: [cmd][addr:4][len:4] and returns len bytes.
    private async Task<byte[]> SpiFlashReadPageAsync(uint addr, uint len, CancellationToken ct)
    {
        if (!Connected || _port == null)
            throw new InvalidOperationException("Device not connected.");

        if (len == 0)
            return Array.Empty<byte>();

        if (len > SpiFlashMaxReadLen)
            len = SpiFlashMaxReadLen;

        var frame = new byte[1 + 4 + 4];
        frame[0] = CmdSpiFlashReadPage;
        BinaryPrimitives.WriteUInt32LittleEndian(frame.AsSpan(1, 4), addr);
        BinaryPrimitives.WriteUInt32LittleEndian(frame.AsSpan(5, 4), len);

        byte[] result = await Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();

            byte[]? rx;
            lock (_port)
            {
                _port!.Open();
                _port!.Write(frame, 0, frame.Length);
                rx = ReadExact((int)len);
                _port!.Close();
            }

            if (rx is null)
                throw new TimeoutException("SPI flash read timed out.");

            return rx;
        }, ct).ConfigureAwait(false);

        return result;
    }

    private byte[]? SpiFlashReadPageNoLock(uint addr, uint len)
    {
        if (!Connected || _port == null)
            throw new InvalidOperationException("Device not connected.");

        if (len == 0)
            return Array.Empty<byte>();

        if (len > SpiFlashMaxReadLen)
            len = SpiFlashMaxReadLen;

        var frame = new byte[1 + 4 + 4];
        frame[0] = CmdSpiFlashReadPage;
        BinaryPrimitives.WriteUInt32LittleEndian(frame.AsSpan(1, 4), addr);
        BinaryPrimitives.WriteUInt32LittleEndian(frame.AsSpan(5, 4), len);

        _port!.Write(frame, 0, frame.Length);
        return ReadExact((int)len);
    }

    /// <summary>
    /// Reads exactly <paramref name="len"/> bytes from SPI flash starting at <paramref name="addr"/>.
    /// Internally performs multiple page reads (<= 256 bytes each).
    /// </summary>
    private async Task<byte[]> SpiFlashReadBytesAsync(
        uint addr,
        uint len,
        IProgress<double>? progress,
        CancellationToken ct)
    {
        if (!Connected || _port == null)
            throw new InvalidOperationException("Device not connected.");

        if (len == 0)
            return Array.Empty<byte>();

        return await Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();

            var result = new byte[len];
            uint read = 0;

            var frame = new byte[1 + 4 + 4];
            frame[0] = CmdSpiFlashReadPage;

            lock (_port)
            {
                _port!.Open();
                _port!.DiscardInBuffer();

                // Pause display updates during large transfers.
                _port!.Write(new byte[] { (byte)UsbCmd.CMD_SCREEN_CHANGE, (byte)SCREEN_CMD.SCREEN_PAUSE_UPDATES }, 0, 2);

                try
                {
                    while (read < len)
                    {
                        uint remaining = len - read;
                        uint toRead = (uint)Math.Min(SpiFlashMaxReadLen, remaining);

                        BinaryPrimitives.WriteUInt32LittleEndian(frame.AsSpan(1, 4), addr + read);
                        BinaryPrimitives.WriteUInt32LittleEndian(frame.AsSpan(5, 4), toRead);

                        _port!.Write(frame, 0, frame.Length);
                        var chunk = ReadExact((int)toRead);

                        if (chunk is null || chunk.Length != toRead)
                            throw new TimeoutException("SPI flash read error.");

                        Buffer.BlockCopy(chunk, 0, result, (int)read, (int)toRead);

                        read += toRead;
                        progress?.Report((double)read / len);
                    }

                    return result;
                }
                finally
                {
                    // Resume display updates.
                    _port!.Write(new byte[] { (byte)UsbCmd.CMD_SCREEN_CHANGE, (byte)SCREEN_CMD.SCREEN_RESUME_UPDATES }, 0, 2);
                    _port!.Close();
                }
            }
        }, ct).ConfigureAwait(false);
    }

    // Firmware erase command takes: [cmd][addr:4][len:4] and returns a 1-byte status (1=OK)
    private bool SpiFlashEraseRangeNoLock(uint addr, uint len)
    {
        var frame = new byte[1 + 4 + 4];
        frame[0] = CmdSpiFlashEraseSector;
        BinaryPrimitives.WriteUInt32LittleEndian(frame.AsSpan(1, 4), addr);
        BinaryPrimitives.WriteUInt32LittleEndian(frame.AsSpan(5, 4), len);

        SpiFlashWriteAllChunked(frame, 64, 0);

        uint timeoutMs = (len / SpiFlashSectorSizeBytes) * 100;

        int result = 0;
        DateTime startTime = DateTime.UtcNow;

        while (result == 0 && DateTime.UtcNow < startTime.AddMilliseconds(timeoutMs)) {
            Thread.Sleep(1);
            byte[] buffer = new byte[1];
            if(_port!.BytesToRead > 0)
            {
                _port!.Read(buffer, 0, 1);
                result = buffer[0];
            } 
        }

        return result == 1;
    }

    private bool SpiFlashWritePageNoLock(uint addr, ReadOnlySpan<byte> data)
    {
        if (data.Length == 0)
            return false;

        if (data.Length > SpiFlashMaxWriteLen)
            throw new ArgumentOutOfRangeException(nameof(data), $"Max write length is {SpiFlashMaxWriteLen} bytes.");

        var header = new byte[1 + 4 + 4];
        header[0] = CmdSpiFlashWritePage;
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(1, 4), addr);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(5, 4), (uint)data.Length);

        SpiFlashWriteAllChunked(header, 64, 0);
        SpiFlashWriteAllChunked(data, 64, 1);

        var result = SpiFlashReadResult((uint)data.Length * 100);
        if (!result)
            throw new InvalidOperationException($"Device reported error on page write at 0x{addr:X8}.");
        return true;
    }

    public Task<byte[]> ReadSpiFlashBytesAsync(
        uint addr,
        uint len,
        IProgress<double>? progress = null,
        CancellationToken ct = default)
        => SpiFlashReadBytesAsync(addr, len, progress, ct);

    public async Task WriteSpiFlashBytesPreserveSectorsAsync(
        uint addr,
        ReadOnlyMemory<byte> data,
        IProgress<double>? progress = null,
        CancellationToken ct = default)
    {
        if (!Connected || _port == null)
            throw new InvalidOperationException("Device not connected.");

        if (data.Length == 0)
            return;

        uint len = (uint)data.Length;
        if (addr + len > SpiFlashLogicalSizeBytes)
            throw new ArgumentOutOfRangeException(nameof(addr), "Write exceeds SPI flash size.");

        await Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();

            uint firstSector = addr / SpiFlashSectorSizeBytes;
            uint firstSectorAddr = firstSector * SpiFlashSectorSizeBytes;
            uint startPageAddr = (uint)firstSector * SpiFlashSectorSizeBytes;
            uint lastSector = (addr + len - 1) / SpiFlashSectorSizeBytes;
            uint lastSectorAddr = lastSector * SpiFlashSectorSizeBytes;
            uint lastPageAddr = lastSector * SpiFlashSectorSizeBytes;
            uint existingDataSize = (lastSectorAddr + SpiFlashSectorSizeBytes) - startPageAddr;

            byte[] writeBuf = new byte[existingDataSize];

            lock (_port)
            {
                _port!.Open();
                _port!.DiscardInBuffer();

                // Pause display updates during large transfers.
                _port!.Write(new byte[] { (byte)UsbCmd.CMD_SCREEN_CHANGE, (byte)SCREEN_CMD.SCREEN_PAUSE_UPDATES }, 0, 2);

                try
                {
                    // Read existing data to existingDataBuf
                    var readFrame = new byte[1 + 4 + 4];
                    readFrame[0] = CmdSpiFlashReadPage;

                    for (uint pageAddr = startPageAddr; pageAddr <= lastPageAddr; pageAddr += SpiFlashPageSize)
                    {
                        ct.ThrowIfCancellationRequested();

                        var pageBuf = SpiFlashReadPageNoLock(pageAddr, SpiFlashPageSize);
                        if(pageBuf is null) {
                            throw new Exception($"SPI flash read error at 0x{pageAddr:X8}.");
                        }
                        Buffer.BlockCopy(pageBuf, 0, writeBuf, (int)(pageAddr - startPageAddr), (int)SpiFlashPageSize);
                        progress?.Report(Math.Clamp((double)(pageAddr - startPageAddr + SpiFlashPageSize) / (3*existingDataSize), 0.1, 1.0));
                    }

                    // Overwrite the portion of writeBuf that will be updated with new data.
                    int offsetIntoExisting = (int)(addr - startPageAddr);
                    data.CopyTo(writeBuf.AsMemory(offsetIntoExisting, (int)len));

                    // Erase existing data sectors
                    SpiFlashEraseRangeNoLock(firstSectorAddr, existingDataSize);
                    progress?.Report(0.66);

                    // Write new data in page-sized chunks
                    for (uint pageAddr = startPageAddr; pageAddr <= lastPageAddr; pageAddr += SpiFlashPageSize)
                    {
                        ct.ThrowIfCancellationRequested();
                        int offsetIntoWriteBuf = (int)(pageAddr - startPageAddr);
                        var pageData = writeBuf.AsSpan(offsetIntoWriteBuf, (int)SpiFlashPageSize);
                        SpiFlashWritePageNoLock(pageAddr, pageData);
                        progress?.Report(Math.Clamp(0.66 + (double)(pageAddr - startPageAddr + SpiFlashPageSize) / (3*existingDataSize), 0.0, 1.0));
                    }
                }
                finally
                {
                    // Resume display updates.
                    _port!.Write(new byte[] { (byte)UsbCmd.CMD_SCREEN_CHANGE, (byte)SCREEN_CMD.SCREEN_RESUME_UPDATES }, 0, 2);
                    _port!.Close();
                    progress?.Report(1.0);
                }
            }
        }, ct).ConfigureAwait(false);
    }

    private bool SpiFlashReadResult(uint timeoutMs)
    {
        int result = 0;
        DateTime startTime = DateTime.UtcNow;

        while (result == 0 && DateTime.UtcNow < startTime.AddMilliseconds(timeoutMs))
        {
            Thread.Sleep(1);
            byte[] buffer = new byte[1];
            if (_port!.BytesToRead > 0)
            {
                _port!.Read(buffer, 0, 1);
                result = buffer[0];
            }
        }
        return result == 1;
    }
}
