using System.Buffers.Binary;

namespace OfficeDocuments.Word.Formatting;

/// <summary>
/// The pixel size and resolution of an image, read from its own header.
/// </summary>
/// <remarks>
/// <para>
/// An inline image needs a size in the document, and the only sensible default is the image's own.
/// Getting that from <c>System.Drawing</c> is not an option — it is Windows-only on .NET Core, and
/// this library multi-targets and runs on Linux in CI — so the few header layouts that matter are
/// read directly. That is a small amount of parsing in exchange for not adding an image dependency
/// to a document library.
/// </para>
/// <para>
/// Only what is needed is parsed: the dimensions and the resolution. Anything not understood returns
/// <see langword="null"/>, and the caller then requires an explicit size rather than guessing.
/// </para>
/// </remarks>
internal sealed record ImageMetadata(
    Enums.ImageType Type,
    int PixelWidth,
    int PixelHeight,
    double HorizontalDpi,
    double VerticalDpi)
{
    /// <summary>
    /// Resolution assumed when a file does not state one. 96 is the conventional screen default and
    /// what Word itself falls back to.
    /// </summary>
    private const double DefaultDpi = 96d;

    /// <summary>
    /// Points per inch.
    /// </summary>
    private const double PointsPerInch = 72d;

    /// <summary>
    /// Metres to inches, for the formats that store resolution in pixels per metre.
    /// </summary>
    private const double MetresPerInch = 0.0254d;

    /// <summary>
    /// Width at the image's own resolution, in points.
    /// </summary>
    internal double WidthInPoints => PixelWidth / HorizontalDpi * PointsPerInch;

    /// <summary>
    /// Height at the image's own resolution, in points.
    /// </summary>
    internal double HeightInPoints => PixelHeight / VerticalDpi * PointsPerInch;

    /// <summary>
    /// Reads the metadata of the image in <paramref name="content"/>, restoring the stream position.
    /// </summary>
    /// <returns>The metadata, or <see langword="null"/> when the format is not recognized.</returns>
    internal static ImageMetadata? TryRead(Stream content)
    {
        if (!content.CanSeek)
        {
            return null;
        }

        var originalPosition = content.Position;

        try
        {
            content.Position = 0;

            var header = new byte[32];
            var headerLength = ReadAtLeast(content, header, header.Length);

            if (StartsWith(header, headerLength, [0x89, (byte)'P', (byte)'N', (byte)'G']))
            {
                return ReadPng(content, header, headerLength);
            }

            if (StartsWith(header, headerLength, [0xFF, 0xD8]))
            {
                return ReadJpeg(content);
            }

            if (StartsWith(header, headerLength, [(byte)'G', (byte)'I', (byte)'F']))
            {
                return ReadGif(header, headerLength);
            }

            if (StartsWith(header, headerLength, [(byte)'B', (byte)'M']))
            {
                return ReadBmp(content);
            }

            return null;
        }
        catch (IOException)
        {
            return null;
        }
        finally
        {
            content.Position = originalPosition;
        }
    }

    /// <summary>
    /// PNG: an 8-byte signature, then an IHDR chunk whose payload starts with width and height as
    /// big-endian 32-bit values. Resolution, if present, is in an optional pHYs chunk.
    /// </summary>
    private static ImageMetadata? ReadPng(Stream content, byte[] header, int headerLength)
    {
        const int ihdrWidthOffset = 16;
        if (headerLength < ihdrWidthOffset + 8)
        {
            return null;
        }

        var width = BinaryPrimitives.ReadInt32BigEndian(header.AsSpan(ihdrWidthOffset, 4));
        var height = BinaryPrimitives.ReadInt32BigEndian(header.AsSpan(ihdrWidthOffset + 4, 4));
        if (width <= 0 || height <= 0)
        {
            return null;
        }

        var (horizontalDpi, verticalDpi) = ReadPngResolution(content);

        return new ImageMetadata(Enums.ImageType.Png, width, height, horizontalDpi, verticalDpi);
    }

    private static (double Horizontal, double Vertical) ReadPngResolution(Stream content)
    {
        // Chunks follow the signature: length (4), type (4), payload, CRC (4).
        content.Position = 8;
        var chunkHeader = new byte[8];

        while (ReadAtLeast(content, chunkHeader, chunkHeader.Length) == chunkHeader.Length)
        {
            var payloadLength = BinaryPrimitives.ReadInt32BigEndian(chunkHeader.AsSpan(0, 4));
            if (payloadLength < 0)
            {
                break;
            }

            var chunkType = System.Text.Encoding.ASCII.GetString(chunkHeader, 4, 4);

            if (string.Equals(chunkType, "pHYs", StringComparison.Ordinal) && payloadLength >= 9)
            {
                var payload = new byte[9];
                if (ReadAtLeast(content, payload, payload.Length) != payload.Length)
                {
                    break;
                }

                // Unit 1 means pixels per metre; anything else leaves the ratio undefined.
                if (payload[8] != 1)
                {
                    break;
                }

                var horizontal = BinaryPrimitives.ReadUInt32BigEndian(payload.AsSpan(0, 4));
                var vertical = BinaryPrimitives.ReadUInt32BigEndian(payload.AsSpan(4, 4));

                return (FromPixelsPerMetre(horizontal), FromPixelsPerMetre(vertical));
            }

            if (string.Equals(chunkType, "IDAT", StringComparison.Ordinal))
            {
                // Pixel data has started, so there is no pHYs chunk to find.
                break;
            }

            content.Position += payloadLength + 4;
        }

        return (DefaultDpi, DefaultDpi);
    }

    /// <summary>
    /// JPEG: a marker stream. The dimensions live in a start-of-frame marker, and the resolution in
    /// the JFIF application marker that usually precedes it.
    /// </summary>
    private static ImageMetadata? ReadJpeg(Stream content)
    {
        content.Position = 2;

        var horizontalDpi = DefaultDpi;
        var verticalDpi = DefaultDpi;
        var marker = new byte[4];

        while (ReadAtLeast(content, marker, 4) == 4)
        {
            if (marker[0] != 0xFF)
            {
                return null;
            }

            var markerType = marker[1];
            var segmentLength = BinaryPrimitives.ReadUInt16BigEndian(marker.AsSpan(2, 2));
            if (segmentLength < 2)
            {
                return null;
            }

            var payloadLength = segmentLength - 2;

            if (markerType == 0xE0 && payloadLength >= 12)
            {
                var payload = new byte[12];
                if (ReadAtLeast(content, payload, payload.Length) != payload.Length)
                {
                    return null;
                }

                // JFIF: identifier (5), version (2), units (1), x density (2), y density (2).
                var units = payload[7];
                var xDensity = BinaryPrimitives.ReadUInt16BigEndian(payload.AsSpan(8, 2));
                var yDensity = BinaryPrimitives.ReadUInt16BigEndian(payload.AsSpan(10, 2));

                if (xDensity > 0 && yDensity > 0)
                {
                    horizontalDpi = units == 2 ? xDensity * 2.54d : xDensity;
                    verticalDpi = units == 2 ? yDensity * 2.54d : yDensity;
                }

                content.Position += payloadLength - payload.Length;
                continue;
            }

            if (IsStartOfFrame(markerType))
            {
                var payload = new byte[5];
                if (ReadAtLeast(content, payload, payload.Length) != payload.Length)
                {
                    return null;
                }

                // Precision (1), height (2), width (2).
                var height = BinaryPrimitives.ReadUInt16BigEndian(payload.AsSpan(1, 2));
                var width = BinaryPrimitives.ReadUInt16BigEndian(payload.AsSpan(3, 2));

                return width > 0 && height > 0
                    ? new ImageMetadata(Enums.ImageType.Jpeg, width, height, horizontalDpi, verticalDpi)
                    : null;
            }

            content.Position += payloadLength;
        }

        return null;
    }

    /// <summary>
    /// A start-of-frame marker carries the dimensions. The arithmetic and Huffman table markers sit in
    /// the same range and must not be mistaken for one.
    /// </summary>
    private static bool IsStartOfFrame(byte markerType)
    {
        return markerType is >= 0xC0 and <= 0xCF
            && markerType is not (0xC4 or 0xC8 or 0xCC);
    }

    /// <summary>
    /// GIF: a 6-byte signature, then width and height as little-endian 16-bit values. The format
    /// stores no resolution.
    /// </summary>
    private static ImageMetadata? ReadGif(byte[] header, int headerLength)
    {
        if (headerLength < 10)
        {
            return null;
        }

        var width = BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(6, 2));
        var height = BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(8, 2));

        return width > 0 && height > 0
            ? new ImageMetadata(Enums.ImageType.Gif, width, height, DefaultDpi, DefaultDpi)
            : null;
    }

    /// <summary>
    /// BMP: a 14-byte file header, then an info header holding the dimensions and, optionally, the
    /// resolution in pixels per metre.
    /// </summary>
    private static ImageMetadata? ReadBmp(Stream content)
    {
        content.Position = 14;

        var infoHeader = new byte[40];
        if (ReadAtLeast(content, infoHeader, infoHeader.Length) != infoHeader.Length)
        {
            return null;
        }

        var width = BinaryPrimitives.ReadInt32LittleEndian(infoHeader.AsSpan(4, 4));

        // A negative height means the rows are stored top-down; the size is the magnitude.
        var height = Math.Abs(BinaryPrimitives.ReadInt32LittleEndian(infoHeader.AsSpan(8, 4)));
        if (width <= 0 || height <= 0)
        {
            return null;
        }

        var horizontalPixelsPerMetre = BinaryPrimitives.ReadInt32LittleEndian(infoHeader.AsSpan(24, 4));
        var verticalPixelsPerMetre = BinaryPrimitives.ReadInt32LittleEndian(infoHeader.AsSpan(28, 4));

        return new ImageMetadata(
            Enums.ImageType.Bmp,
            width,
            height,
            horizontalPixelsPerMetre > 0 ? FromPixelsPerMetre((uint)horizontalPixelsPerMetre) : DefaultDpi,
            verticalPixelsPerMetre > 0 ? FromPixelsPerMetre((uint)verticalPixelsPerMetre) : DefaultDpi);
    }

    private static double FromPixelsPerMetre(uint pixelsPerMetre)
    {
        return pixelsPerMetre > 0 ? pixelsPerMetre * MetresPerInch : DefaultDpi;
    }

    private static bool StartsWith(byte[] buffer, int length, byte[] signature)
    {
        if (length < signature.Length)
        {
            return false;
        }

        for (var index = 0; index < signature.Length; index++)
        {
            if (buffer[index] != signature[index])
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Fills as much of <paramref name="buffer"/> as the stream provides, since a single
    /// <see cref="Stream.Read(byte[], int, int)"/> is not required to return everything asked for.
    /// </summary>
    private static int ReadAtLeast(Stream content, byte[] buffer, int count)
    {
        var total = 0;
        while (total < count)
        {
            var read = content.Read(buffer, total, count - total);
            if (read == 0)
            {
                break;
            }

            total += read;
        }

        return total;
    }
}
