using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;
using OfficeDocuments.Word.TestKit;

namespace OfficeDocuments.Word.Tests;

/// <summary>
/// Direct tests of the image header reader.
/// </summary>
/// <remarks>
/// Tested here rather than only through <see cref="ImageTest"/> because the reader is pure logic over
/// four different binary layouts, and the document tests only ever exercise the PNG path. A malformed
/// read shows up as a wrongly sized image, which is exactly the kind of defect that is hard to notice
/// in a rendered document.
/// </remarks>
public class ImageMetadataTest
{
    [Fact]
    public void TryRead_Png_ReadsTypeAndDimensions()
    {
        using var content = new MemoryStream(TestImages.PngWithSize(640, 480));

        var metadata = ImageMetadata.TryRead(content);

        Assert.NotNull(metadata);
        Assert.Equal(ImageType.Png, metadata.Type);
        Assert.Equal(640, metadata.PixelWidth);
        Assert.Equal(480, metadata.PixelHeight);
    }

    /// <summary>
    /// A PNG without a <c>pHYs</c> chunk states no resolution, and 96 DPI is the conventional default.
    /// </summary>
    [Fact]
    public void TryRead_PngWithoutResolution_AssumesNinetySixDpi()
    {
        using var content = new MemoryStream(TestImages.PngWithSize(96, 96));

        var metadata = ImageMetadata.TryRead(content);

        Assert.NotNull(metadata);
        Assert.Equal(96d, metadata.HorizontalDpi, precision: 1);
        Assert.Equal(72d, metadata.WidthInPoints, precision: 1);
    }

    [Fact]
    public void TryRead_PngWithResolution_UsesIt()
    {
        using var content = new MemoryStream(TestImages.PngWithSize(300, 300, dotsPerInch: 150));

        var metadata = ImageMetadata.TryRead(content);

        Assert.NotNull(metadata);
        Assert.Equal(150d, metadata.HorizontalDpi, precision: 1);
        Assert.Equal(144d, metadata.WidthInPoints, precision: 1);
    }

    [Fact]
    public void TryRead_Jpeg_ReadsTypeDimensionsAndResolution()
    {
        using var content = new MemoryStream(BuildJpeg(width: 200, height: 100, densityPerInch: 150));

        var metadata = ImageMetadata.TryRead(content);

        Assert.NotNull(metadata);
        Assert.Equal(ImageType.Jpeg, metadata.Type);
        Assert.Equal(200, metadata.PixelWidth);
        Assert.Equal(100, metadata.PixelHeight);
        Assert.Equal(150d, metadata.HorizontalDpi, precision: 1);
    }

    /// <summary>
    /// The Huffman table marker sits in the same numeric range as the start-of-frame markers. Reading it
    /// as one would take the table's bytes for a size.
    /// </summary>
    [Fact]
    public void TryRead_JpegWithHuffmanTable_SkipsItAndFindsTheRealFrame()
    {
        using var content = new MemoryStream(BuildJpeg(width: 320, height: 240, densityPerInch: 96, includeHuffmanTable: true));

        var metadata = ImageMetadata.TryRead(content);

        Assert.NotNull(metadata);
        Assert.Equal(320, metadata.PixelWidth);
        Assert.Equal(240, metadata.PixelHeight);
    }

    [Fact]
    public void TryRead_Gif_ReadsTypeAndDimensions()
    {
        using var content = new MemoryStream(BuildGif(width: 320, height: 200));

        var metadata = ImageMetadata.TryRead(content);

        Assert.NotNull(metadata);
        Assert.Equal(ImageType.Gif, metadata.Type);
        Assert.Equal(320, metadata.PixelWidth);
        Assert.Equal(200, metadata.PixelHeight);
    }

    [Fact]
    public void TryRead_Bmp_ReadsTypeDimensionsAndResolution()
    {
        using var content = new MemoryStream(BuildBmp(width: 150, height: 75, dotsPerInch: 150));

        var metadata = ImageMetadata.TryRead(content);

        Assert.NotNull(metadata);
        Assert.Equal(ImageType.Bmp, metadata.Type);
        Assert.Equal(150, metadata.PixelWidth);
        Assert.Equal(75, metadata.PixelHeight);
        Assert.Equal(150d, metadata.HorizontalDpi, precision: 0);
    }

    /// <summary>
    /// A negative height means the rows are stored top-down, not that the image has a negative size.
    /// </summary>
    [Fact]
    public void TryRead_BmpStoredTopDown_ReadsThePositiveHeight()
    {
        using var content = new MemoryStream(BuildBmp(width: 10, height: -20, dotsPerInch: 96));

        var metadata = ImageMetadata.TryRead(content);

        Assert.NotNull(metadata);
        Assert.Equal(20, metadata.PixelHeight);
    }

    [Fact]
    public void TryRead_UnrecognizedContent_ReturnsNull()
    {
        using var content = new MemoryStream(TestImages.UnrecognizedImage());

        Assert.Null(ImageMetadata.TryRead(content));
    }

    [Fact]
    public void TryRead_EmptyContent_ReturnsNull()
    {
        using var content = new MemoryStream();

        Assert.Null(ImageMetadata.TryRead(content));
    }

    /// <summary>
    /// Nothing can be read from a stream that cannot rewind, and the caller has to supply the size.
    /// </summary>
    [Fact]
    public void TryRead_NonSeekableStream_ReturnsNull()
    {
        using var content = new NonSeekableStream(TestImages.MinimalPng());

        Assert.Null(ImageMetadata.TryRead(content));
    }

    /// <summary>
    /// The reader must leave the stream where it found it, or the caller writes a truncated image part.
    /// </summary>
    [Fact]
    public void TryRead_RestoresTheStreamPosition()
    {
        using var content = new MemoryStream(TestImages.PngWithSize(10, 10));
        content.Position = 4;

        ImageMetadata.TryRead(content);

        Assert.Equal(4, content.Position);
    }

    private static byte[] BuildJpeg(int width, int height, int densityPerInch, bool includeHuffmanTable = false)
    {
        var bytes = new List<byte> { 0xFF, 0xD8 };

        // APP0 with a JFIF header: length, identifier, version, units, densities, thumbnail size.
        bytes.AddRange([0xFF, 0xE0, 0x00, 0x10]);
        bytes.AddRange("JFIF"u8);
        bytes.Add(0x00);
        bytes.AddRange([0x01, 0x01, 0x01]);
        bytes.AddRange(BigEndian16(densityPerInch));
        bytes.AddRange(BigEndian16(densityPerInch));
        bytes.AddRange([0x00, 0x00]);

        if (includeHuffmanTable)
        {
            bytes.AddRange([0xFF, 0xC4, 0x00, 0x05, 0x00, 0x00, 0x00]);
        }

        // SOF0: length, sample precision, height, width, component count.
        bytes.AddRange([0xFF, 0xC0, 0x00, 0x0B, 0x08]);
        bytes.AddRange(BigEndian16(height));
        bytes.AddRange(BigEndian16(width));
        bytes.AddRange([0x01, 0x01, 0x11, 0x00]);

        return [.. bytes];
    }

    private static byte[] BuildGif(int width, int height)
    {
        var bytes = new List<byte>();
        bytes.AddRange("GIF89a"u8);
        bytes.AddRange(LittleEndian16(width));
        bytes.AddRange(LittleEndian16(height));
        bytes.AddRange([0x00, 0x00, 0x00]);

        return [.. bytes];
    }

    private static byte[] BuildBmp(int width, int height, double dotsPerInch)
    {
        const double metresPerInch = 0.0254d;
        var pixelsPerMetre = (int)Math.Round(dotsPerInch / metresPerInch);

        var bytes = new List<byte>();
        bytes.AddRange("BM"u8);
        bytes.AddRange(new byte[12]);

        // BITMAPINFOHEADER: size, width, height, planes, bit count, compression, image size,
        // horizontal and vertical resolution, palette counts.
        bytes.AddRange(LittleEndian32(40));
        bytes.AddRange(LittleEndian32(width));
        bytes.AddRange(LittleEndian32(height));
        bytes.AddRange(LittleEndian16(1));
        bytes.AddRange(LittleEndian16(24));
        bytes.AddRange(LittleEndian32(0));
        bytes.AddRange(LittleEndian32(0));
        bytes.AddRange(LittleEndian32(pixelsPerMetre));
        bytes.AddRange(LittleEndian32(pixelsPerMetre));
        bytes.AddRange(LittleEndian32(0));
        bytes.AddRange(LittleEndian32(0));

        return [.. bytes];
    }

    private static byte[] BigEndian16(int value) => [(byte)(value >> 8), (byte)value];

    private static byte[] LittleEndian16(int value) => [(byte)value, (byte)(value >> 8)];

    private static byte[] LittleEndian32(int value) =>
        [(byte)value, (byte)(value >> 8), (byte)(value >> 16), (byte)(value >> 24)];

    /// <summary>
    /// A forward-only stream, standing in for a network or compressed source.
    /// </summary>
    private sealed class NonSeekableStream(byte[] content) : MemoryStream(content)
    {
        public override bool CanSeek => false;
    }
}
