namespace OfficeDocuments.Word.TestKit;

/// <summary>
/// Image fixtures, so image tests need no binary files in the repository.
/// </summary>
public static class TestImages
{
    /// <summary>
    /// A valid 1×1 PNG.
    /// </summary>
    /// <remarks>
    /// Real bytes rather than a stub, because the document has to embed something a reader will accept.
    /// The size makes it useless for checking rendered dimensions, which is what
    /// <see cref="PngWithSize"/> is for.
    /// </remarks>
    public static byte[] MinimalPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=");

    /// <summary>
    /// A PNG whose header declares <paramref name="pixelWidth"/> by <paramref name="pixelHeight"/>,
    /// optionally with a resolution.
    /// </summary>
    /// <remarks>
    /// The pixel data still describes a 1×1 image, so this is a fixture for the header-reading and
    /// sizing paths only — not something to render. It exists because testing that a 400-pixel image
    /// lands at the right point size needs a 400-pixel header, and shipping such a file as a binary
    /// fixture would hide what the test depends on.
    /// </remarks>
    /// <param name="pixelWidth">Width to declare.</param>
    /// <param name="pixelHeight">Height to declare.</param>
    /// <param name="dotsPerInch">Resolution to declare, or <see langword="null"/> to omit it.</param>
    public static byte[] PngWithSize(int pixelWidth, int pixelHeight, double? dotsPerInch = null)
    {
        var png = MinimalPng().ToList();

        WriteBigEndian(png, 16, pixelWidth);
        WriteBigEndian(png, 20, pixelHeight);
        FixChunkCrc(png, chunkStart: 8);

        if (dotsPerInch is not null)
        {
            InsertPhysicalDimensions(png, dotsPerInch.Value);
        }

        return [.. png];
    }

    /// <summary>
    /// Bytes that are not any recognized image format.
    /// </summary>
    public static byte[] UnrecognizedImage() => [0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07];

    /// <summary>
    /// Inserts a <c>pHYs</c> chunk declaring a resolution, immediately after <c>IHDR</c>.
    /// </summary>
    private static void InsertPhysicalDimensions(List<byte> png, double dotsPerInch)
    {
        const double metresPerInch = 0.0254d;
        var pixelsPerMetre = (int)Math.Round(dotsPerInch / metresPerInch);

        var chunk = new List<byte>();
        chunk.AddRange([0x00, 0x00, 0x00, 0x09]);
        chunk.AddRange("pHYs"u8);
        chunk.AddRange(ToBigEndian(pixelsPerMetre));
        chunk.AddRange(ToBigEndian(pixelsPerMetre));
        chunk.Add(0x01);
        chunk.AddRange([0x00, 0x00, 0x00, 0x00]);

        // IHDR is the first chunk: 8-byte signature, then 4 length + 4 type + 13 payload + 4 CRC.
        const int afterIhdr = 8 + 4 + 4 + 13 + 4;
        png.InsertRange(afterIhdr, chunk);
        FixChunkCrc(png, afterIhdr);
    }

    private static void WriteBigEndian(List<byte> buffer, int offset, int value)
    {
        var bytes = ToBigEndian(value);
        for (var index = 0; index < bytes.Length; index++)
        {
            buffer[offset + index] = bytes[index];
        }
    }

    private static byte[] ToBigEndian(int value) =>
    [
        (byte)(value >> 24),
        (byte)(value >> 16),
        (byte)(value >> 8),
        (byte)value,
    ];

    /// <summary>
    /// Recomputes a chunk's CRC after its payload changed, so the file stays a valid PNG.
    /// </summary>
    private static void FixChunkCrc(List<byte> png, int chunkStart)
    {
        var payloadLength = (png[chunkStart] << 24)
                            | (png[chunkStart + 1] << 16)
                            | (png[chunkStart + 2] << 8)
                            | png[chunkStart + 3];

        // The CRC covers the type and the payload, but not the length.
        var crcStart = chunkStart + 4;
        var crcLength = 4 + payloadLength;
        var crc = Crc32(png, crcStart, crcLength);

        WriteBigEndian(png, crcStart + crcLength, unchecked((int)crc));
    }

    private static uint Crc32(List<byte> buffer, int offset, int count)
    {
        var crc = 0xFFFFFFFFU;

        for (var index = offset; index < offset + count; index++)
        {
            crc ^= buffer[index];
            for (var bit = 0; bit < 8; bit++)
            {
                crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320U : crc >> 1;
            }
        }

        return crc ^ 0xFFFFFFFFU;
    }
}
