using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.Tests;

/// <summary>
/// Covers the document-level page setup and core properties.
/// </summary>
public class PageSetupAndMetadataTest : WordTestBase
{
    [Fact]
    public void ApplyPageSetup_RoundTripsSizeOrientationAndMargins()
    {
        var setup = new PageSetup
        {
            PaperSize = PaperSize.A4,
            Orientation = PageOrientation.Portrait,
            MarginTop = 56,
            MarginBottom = 56,
            MarginLeft = 70,
            MarginRight = 42,
            HeaderDistance = 28,
            FooterDistance = 28,
        };

        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            document.GetBody().AddParagraph("Body");
            document.ApplyPageSetup(setup);
        }

        OpenXmlValidation.AssertValid(stream);

        using var reopened = OpenDocument(stream, isEditable: false);
        var actual = reopened.PageSetup;

        Assert.Equal(PaperSize.A4, actual.PaperSize);
        Assert.Equal(PageOrientation.Portrait, actual.Orientation);
        Assert.Equal(56, actual.MarginTop);
        Assert.Equal(56, actual.MarginBottom);
        Assert.Equal(70, actual.MarginLeft);
        Assert.Equal(42, actual.MarginRight);
        Assert.Equal(28, actual.HeaderDistance);
        Assert.Equal(28, actual.FooterDistance);
    }

    /// <summary>
    /// Landscape is not just an attribute: the stored width and height have to be swapped as well, or
    /// Word lays the text out on a portrait page and only the printer setting changes.
    /// </summary>
    [Fact]
    public void ApplyPageSetup_Landscape_SwapsTheStoredDimensions()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            document.GetBody().AddParagraph("Body");
            document.ApplyPageSetup(new PageSetup { PaperSize = PaperSize.A4, Orientation = PageOrientation.Landscape });
        }

        OpenXmlValidation.AssertValid(stream);

        var size = ReadDocumentElement(
            stream,
            document => document.Body!.GetFirstChild<WordLib.SectionProperties>()!.GetFirstChild<WordLib.PageSize>()!);

        Assert.True(size.Width!.Value > size.Height!.Value, "Landscape must store the wider dimension as the width.");
        Assert.Equal(WordLib.PageOrientationValues.Landscape, size.Orient?.Value);
    }

    [Fact]
    public void ApplyPageSetup_WithCustomSize_UsesTheGivenDimensions()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            document.GetBody().AddParagraph("Body");
            document.ApplyPageSetup(new PageSetup { PageWidth = 400, PageHeight = 600 });
        }

        OpenXmlValidation.AssertValid(stream);

        using var reopened = OpenDocument(stream, isEditable: false);
        var actual = reopened.PageSetup;

        Assert.Equal(400, actual.PageWidth);
        Assert.Equal(600, actual.PageHeight);
        Assert.Null(actual.PaperSize);
    }

    /// <summary>
    /// Points in, twips out: 56 points is 1120 twentieths of a point.
    /// </summary>
    [Fact]
    public void ApplyPageSetup_WritesMarginsInTwips()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            document.GetBody().AddParagraph("Body");
            document.ApplyPageSetup(new PageSetup().WithUniformMargins(56));
        }

        var margins = ReadDocumentElement(
            stream,
            document => document.Body!.GetFirstChild<WordLib.SectionProperties>()!.GetFirstChild<WordLib.PageMargin>()!);

        Assert.Equal(1120, margins.Top!.Value);
        Assert.Equal(1120U, margins.Left!.Value);
    }

    /// <summary>
    /// Setting one margin must not zero the others, so the element is seeded with Word's defaults.
    /// </summary>
    [Fact]
    public void ApplyPageSetup_WithOneMargin_LeavesTheOthersAtWordsDefaults()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            document.GetBody().AddParagraph("Body");
            document.ApplyPageSetup(new PageSetup { MarginTop = 100 });
        }

        using var reopened = OpenDocument(stream, isEditable: false);
        var actual = reopened.PageSetup;

        Assert.Equal(100, actual.MarginTop);
        Assert.Equal(72, actual.MarginLeft);
        Assert.Equal(72, actual.MarginRight);
        Assert.Equal(72, actual.MarginBottom);
    }

    [Fact]
    public void ApplyPageSetup_AppliedTwice_KeepsWhatTheSecondCallDoesNotSet()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        document.GetBody().AddParagraph("Body");

        document.ApplyPageSetup(new PageSetup { PaperSize = PaperSize.Letter });
        document.ApplyPageSetup(new PageSetup { MarginLeft = 90 });

        var actual = document.PageSetup;

        Assert.Equal(PaperSize.Letter, actual.PaperSize);
        Assert.Equal(90, actual.MarginLeft);
    }

    [Fact]
    public void SetMetadata_RoundTripsEveryProperty()
    {
        var created = new DateTimeOffset(2026, 7, 27, 9, 30, 0, TimeSpan.Zero);

        var metadata = new DocumentMetadata
        {
            Title = "Quarterly report",
            Subject = "Finance",
            Author = "Martin Domanský",
            Keywords = "report;finance;Q3",
            Description = "Unaudited figures.",
            Category = "Report",
            LastModifiedBy = "Automation",
            Created = created,
        };

        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            document.GetBody().AddParagraph("Body");
            document.SetMetadata(metadata);
        }

        OpenXmlValidation.AssertValid(stream);

        using var reopened = OpenDocument(stream, isEditable: false);
        var actual = reopened.Metadata;

        Assert.Equal("Quarterly report", actual.Title);
        Assert.Equal("Finance", actual.Subject);
        Assert.Equal("Martin Domanský", actual.Author);
        Assert.Equal("report;finance;Q3", actual.Keywords);
        Assert.Equal("Unaudited figures.", actual.Description);
        Assert.Equal("Report", actual.Category);
        Assert.Equal("Automation", actual.LastModifiedBy);
        Assert.Equal(created.UtcDateTime, actual.Created?.UtcDateTime);
    }

    /// <summary>
    /// Metadata goes into the package's core properties, which is what a document management system or
    /// a search index reads — not into the document body.
    /// </summary>
    [Fact]
    public void SetMetadata_WritesToThePackageCoreProperties()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            document.GetBody().AddParagraph("Body");
            document.SetMetadata(new DocumentMetadata { Title = "From the core properties" });
        }

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);

        Assert.Equal("From the core properties", package.PackageProperties.Title);
    }

    [Fact]
    public void SetMetadata_LeavesUnsetPropertiesAlone()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        document.GetBody().AddParagraph("Body");

        document.SetMetadata(new DocumentMetadata { Title = "Original", Author = "First author" });
        document.SetMetadata(new DocumentMetadata { Title = "Replaced" });

        Assert.Equal("Replaced", document.Metadata.Title);
        Assert.Equal("First author", document.Metadata.Author);
    }

    [Fact]
    public void SetMetadata_WithEmptyString_ClearsTheProperty()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        document.GetBody().AddParagraph("Body");

        document.SetMetadata(new DocumentMetadata { Title = "Original" });
        document.SetMetadata(new DocumentMetadata { Title = string.Empty });

        Assert.Equal(string.Empty, document.Metadata.Title);
    }

    [Fact]
    public void ApplyPageSetup_AfterClose_Throws()
    {
        using var stream = new MemoryStream();
        var document = CreateDocument(stream);
        document.GetBody().AddParagraph("Body");
        document.Close();

        Assert.Throws<ObjectDisposedException>(() => document.ApplyPageSetup(new PageSetup()));
    }

    [Theory]
    [InlineData(0d)]
    [InlineData(-10d)]
    public void PageWidth_NotPositive_ThrowsOnAssignment(double width)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new PageSetup { PageWidth = width });
    }

    [Fact]
    public void MarginLeft_Negative_ThrowsOnAssignment()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new PageSetup { MarginLeft = -1 });
    }

    [Fact]
    public void WithUniformMargins_SetsAllFour()
    {
        var setup = new PageSetup { PaperSize = PaperSize.A4 }.WithUniformMargins(50);

        Assert.Equal(50, setup.MarginTop);
        Assert.Equal(50, setup.MarginBottom);
        Assert.Equal(50, setup.MarginLeft);
        Assert.Equal(50, setup.MarginRight);
        Assert.Equal(PaperSize.A4, setup.PaperSize);
    }
}
