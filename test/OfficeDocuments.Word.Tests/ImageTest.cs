using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;
using OfficeDocuments.Word.TestKit;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Pictures = DocumentFormat.OpenXml.Drawing.Pictures;
using WordDrawing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeDocuments.Word.Tests;

/// <summary>
/// Covers inline images: the media part, the relationship, the drawing structure, and sizing.
/// </summary>
public class ImageTest : WordTestBase
{
    /// <summary>
    /// English Metric Units per point, the unit DrawingML measures in.
    /// </summary>
    private const double EmuPerPoint = 12700d;

    /// <summary>
    /// The expected extent, in the units the document stores.
    /// </summary>
    private static long Emu(double points) => (long)Math.Round(points * EmuPerPoint);

    [Fact]
    public void AddImage_FromStream_AddsAnImagePartAndReferencesIt()
    {
        using var stream = WriteAndValidate(body =>
        {
            using var image = new MemoryStream(TestImages.MinimalPng());
            body.AddParagraph().AddImage(image);
        });

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);
        var imagePart = Assert.Single(package.MainDocumentPart!.ImageParts);
        var relationshipId = package.MainDocumentPart.GetIdOfPart(imagePart);
        var blip = package.MainDocumentPart.Document!.Descendants<Drawing.Blip>().Single();

        Assert.Equal(relationshipId, blip.Embed?.Value);
    }

    [Fact]
    public void AddImage_WritesTheEmbeddedImageBytesUnchanged()
    {
        var expected = TestImages.MinimalPng();

        using var stream = WriteAndValidate(body =>
        {
            using var image = new MemoryStream(expected);
            body.AddParagraph().AddImage(image);
        });

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);
        using var embedded = package.MainDocumentPart!.ImageParts.Single().GetStream();
        using var buffer = new MemoryStream();
        embedded.CopyTo(buffer);

        Assert.Equal(expected, buffer.ToArray());
    }

    /// <summary>
    /// The whole point of the builder is that Word accepts the result, and the drawing shape is where
    /// that is easiest to get wrong. The schema gate in <see cref="WordTestBase.WriteAndValidate"/>
    /// checks validity; this checks the pieces are actually wired together.
    /// </summary>
    [Fact]
    public void AddImage_BuildsTheFullInlineDrawingStructure()
    {
        using var stream = WriteAndValidate(body =>
        {
            using var image = new MemoryStream(TestImages.MinimalPng());
            body.AddParagraph().AddImage(image);
        });

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);
        var document = package.MainDocumentPart!.Document!;

        Assert.Single(document.Descendants<WordDrawing.Inline>());
        Assert.Single(document.Descendants<WordDrawing.Extent>());
        Assert.Single(document.Descendants<WordDrawing.DocProperties>());
        Assert.Single(document.Descendants<Drawing.Graphic>());
        Assert.Single(document.Descendants<Pictures.Picture>());
        Assert.Single(document.Descendants<Drawing.Transform2D>());
    }

    /// <summary>
    /// A 400×200 image at 96 DPI is 300×150 points. Deriving that from the file rather than asking the
    /// caller is what makes an unsized image come out at its natural size.
    /// </summary>
    [Fact]
    public void AddImage_WithoutSize_UsesTheImagesOwnDimensions()
    {
        using var stream = WriteAndValidate(body =>
        {
            using var image = new MemoryStream(TestImages.PngWithSize(400, 200));
            body.AddParagraph().AddImage(image);
        });

        var extent = ReadDocumentElement(stream, document => document.Descendants<WordDrawing.Extent>().Single());

        Assert.Equal(Emu(300), extent.Cx!.Value);
        Assert.Equal(Emu(150), extent.Cy!.Value);
    }

    /// <summary>
    /// A file that states its resolution should be honoured: 600 pixels at 300 DPI is 2 inches, so
    /// 144 points, not the 450 points a 96 DPI assumption would give.
    /// </summary>
    /// <remarks>
    /// Compared to one decimal place rather than exactly, because PNG stores resolution as a whole
    /// number of pixels per metre and 300 DPI is 11811.02… of them. The value is therefore not exactly
    /// recoverable from any PNG, and demanding an exact match would be testing the format's rounding
    /// rather than this library's arithmetic.
    /// </remarks>
    [Fact]
    public void AddImage_WithoutSize_HonoursTheImagesStatedResolution()
    {
        using var stream = WriteAndValidate(body =>
        {
            using var image = new MemoryStream(TestImages.PngWithSize(600, 300, dotsPerInch: 300));
            body.AddParagraph().AddImage(image);
        });

        var extent = ReadDocumentElement(stream, document => document.Descendants<WordDrawing.Extent>().Single());

        Assert.Equal(144d, extent.Cx!.Value / EmuPerPoint, precision: 1);
        Assert.Equal(72d, extent.Cy!.Value / EmuPerPoint, precision: 1);
    }

    [Fact]
    public void AddImage_WithExactSize_UsesItVerbatim()
    {
        using var stream = WriteAndValidate(body =>
        {
            using var image = new MemoryStream(TestImages.PngWithSize(400, 200));
            body.AddParagraph().AddImage(image, ImageSize.Exact(120, 90));
        });

        var extent = ReadDocumentElement(stream, document => document.Descendants<WordDrawing.Extent>().Single());

        Assert.Equal(Emu(120), extent.Cx!.Value);
        Assert.Equal(Emu(90), extent.Cy!.Value);
    }

    [Fact]
    public void AddImage_WithWidthOnly_DerivesTheHeightFromTheAspectRatio()
    {
        using var stream = WriteAndValidate(body =>
        {
            using var image = new MemoryStream(TestImages.PngWithSize(400, 200));
            body.AddParagraph().AddImage(image, ImageSize.FromWidth(200));
        });

        var extent = ReadDocumentElement(stream, document => document.Descendants<WordDrawing.Extent>().Single());

        Assert.Equal(Emu(200), extent.Cx!.Value);
        Assert.Equal(Emu(100), extent.Cy!.Value);
    }

    [Fact]
    public void AddImage_WithHeightOnly_DerivesTheWidthFromTheAspectRatio()
    {
        using var stream = WriteAndValidate(body =>
        {
            using var image = new MemoryStream(TestImages.PngWithSize(400, 200));
            body.AddParagraph().AddImage(image, ImageSize.FromHeight(50));
        });

        var extent = ReadDocumentElement(stream, document => document.Descendants<WordDrawing.Extent>().Single());

        Assert.Equal(Emu(100), extent.Cx!.Value);
        Assert.Equal(Emu(50), extent.Cy!.Value);
    }

    /// <summary>
    /// The rendered size and the shape's own extents both describe the picture and have to agree, or
    /// Word shows the image cropped or scaled unexpectedly.
    /// </summary>
    [Fact]
    public void AddImage_SizesTheInlineExtentAndTheShapeExtentsAlike()
    {
        using var stream = WriteAndValidate(body =>
        {
            using var image = new MemoryStream(TestImages.PngWithSize(400, 200));
            body.AddParagraph().AddImage(image, ImageSize.Exact(60, 30));
        });

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);
        var document = package.MainDocumentPart!.Document!;
        var inlineExtent = document.Descendants<WordDrawing.Extent>().Single();
        var shapeExtents = document.Descendants<Drawing.Extents>().Single();

        Assert.Equal(inlineExtent.Cx?.Value, shapeExtents.Cx?.Value);
        Assert.Equal(inlineExtent.Cy?.Value, shapeExtents.Cy?.Value);
    }

    [Fact]
    public void AddImage_FromFilePath_InfersTheTypeFromTheExtension()
    {
        var filePath = GetFilepath("logo.png");
        File.WriteAllBytes(filePath, TestImages.PngWithSize(100, 50));

        using var stream = WriteAndValidate(body => body.AddParagraph().AddImage(filePath));

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);

        Assert.Single(package.MainDocumentPart!.ImageParts);
        Assert.Equal("image/png", package.MainDocumentPart.ImageParts.Single().ContentType);
    }

    /// <summary>
    /// The file name is a better default than "Picture 1" because it is what Word shows in the
    /// selection pane.
    /// </summary>
    [Fact]
    public void AddImage_FromFilePath_NamesTheDrawingAfterTheFile()
    {
        var filePath = GetFilepath("company-logo.png");
        File.WriteAllBytes(filePath, TestImages.MinimalPng());

        using var stream = WriteAndValidate(body => body.AddParagraph().AddImage(filePath));

        var properties = ReadDocumentElement(stream, document => document.Descendants<WordDrawing.DocProperties>().Single());

        Assert.Equal("company-logo.png", properties.Name?.Value);
    }

    [Fact]
    public void AddImage_WithDescription_WritesItAsAlternativeText()
    {
        using var stream = WriteAndValidate(body =>
        {
            using var image = new MemoryStream(TestImages.MinimalPng());
            body.AddParagraph().AddImage(image, size: null, description: "Company logo");
        });

        var properties = ReadDocumentElement(stream, document => document.Descendants<WordDrawing.DocProperties>().Single());

        Assert.Equal("Company logo", properties.Description?.Value);
    }

    /// <summary>
    /// Drawing identifiers have to be unique within the document part.
    /// </summary>
    [Fact]
    public void AddImage_Repeatedly_GivesEachDrawingADistinctId()
    {
        using var stream = WriteAndValidate(body =>
        {
            for (var index = 0; index < 3; index++)
            {
                using var image = new MemoryStream(TestImages.MinimalPng());
                body.AddParagraph().AddImage(image);
            }
        });

        var ids = ReadDocumentElement(
            stream,
            document => document.Descendants<WordDrawing.DocProperties>().Select(properties => properties.Id?.Value).ToList());

        Assert.Equal(3, ids.Count);
        Assert.Equal(3, ids.Distinct().Count());
    }

    /// <summary>
    /// Appending to a document that already contains a drawing must not reuse its identifier.
    /// </summary>
    [Fact]
    public void AddImage_AfterReopening_DoesNotReuseAnExistingDrawingId()
    {
        var filePath = GetFilepath("reopen-image.docx");

        using (var document = CreateDocument(filePath))
        {
            using var image = new MemoryStream(TestImages.MinimalPng());
            document.GetBody().AddParagraph().AddImage(image);
        }

        using (var document = OpenDocument(filePath))
        {
            using var image = new MemoryStream(TestImages.MinimalPng());
            document.GetBody().AddParagraph().AddImage(image);
        }

        OpenXmlValidation.AssertValid(filePath);

        using var package = WordprocessingDocument.Open(filePath, false);
        var ids = package.MainDocumentPart!.Document!
            .Descendants<WordDrawing.DocProperties>()
            .Select(properties => properties.Id?.Value)
            .ToList();

        Assert.Equal(2, ids.Count);
        Assert.Equal(2, ids.Distinct().Count());
    }

    [Fact]
    public void AddImage_InATableCell_IsSchemaValid()
    {
        using var stream = WriteAndValidate(body =>
        {
            using var image = new MemoryStream(TestImages.MinimalPng());
            body.AddTable(1, 1).GetCell(0, 0).Paragraphs[0].AddImage(image);
        });

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);

        Assert.Single(package.MainDocumentPart!.ImageParts);
    }

    [Fact]
    public void AddImage_WithUnrecognizedContent_ThrowsAskingForTheType()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var paragraph = document.GetBody().AddParagraph();
        using var image = new MemoryStream(TestImages.UnrecognizedImage());

        var exception = Assert.Throws<ArgumentException>(() => paragraph.AddImage(image));

        Assert.Contains("image type", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// An explicit type is enough to embed content whose header cannot be read, as long as the caller
    /// also supplies the size the library can no longer infer.
    /// </summary>
    [Fact]
    public void AddImage_WithExplicitTypeAndSize_AcceptsUnreadableContent()
    {
        using var stream = WriteAndValidate(body =>
        {
            using var image = new MemoryStream(TestImages.UnrecognizedImage());
            body.AddParagraph().AddImage(image, ImageType.Png, ImageSize.Exact(72, 72));
        });

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);

        Assert.Single(package.MainDocumentPart!.ImageParts);
    }

    [Fact]
    public void AddImage_WithExplicitTypeButNoReadableSize_ThrowsAskingForAnExactSize()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var paragraph = document.GetBody().AddParagraph();
        using var image = new MemoryStream(TestImages.UnrecognizedImage());

        var exception = Assert.Throws<ArgumentException>(() => paragraph.AddImage(image, ImageType.Png));

        Assert.Contains("Exact", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AddImage_FromFilePathWithUnsupportedExtension_Throws()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var paragraph = document.GetBody().AddParagraph();

        Assert.Throws<ArgumentException>(() => paragraph.AddImage("logo.svg"));
    }

    [Theory]
    [InlineData(0d)]
    [InlineData(-5d)]
    [InlineData(double.NaN)]
    public void ImageSize_WithInvalidDimension_Throws(double dimension)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => ImageSize.FromWidth(dimension));
        Assert.Throws<ArgumentOutOfRangeException>(() => ImageSize.Exact(dimension, 10));
    }
}
