using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;
using OfficeDocuments.Word.TestKit;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.Tests;

/// <summary>
/// Covers headers and footers: their parts, their references, and the switches they depend on.
/// </summary>
public class HeaderFooterTest : WordTestBase
{
    [Fact]
    public void AddHeader_CreatesAPartAndReferencesItFromTheSection()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            document.GetBody().AddParagraph("Body text");
            document.AddHeader().AddParagraph("Page header");
        }

        OpenXmlValidation.AssertValid(stream);

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);
        var headerPart = Assert.Single(package.MainDocumentPart!.HeaderParts);
        var reference = package.MainDocumentPart.Document!.Descendants<WordLib.HeaderReference>().Single();

        Assert.Equal(package.MainDocumentPart.GetIdOfPart(headerPart), reference.Id?.Value);
        Assert.Equal("Page header", headerPart.Header!.InnerText);
    }

    [Fact]
    public void AddFooter_CreatesAPartAndReferencesItFromTheSection()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            document.GetBody().AddParagraph("Body text");
            document.AddFooter().AddParagraph("Page footer");
        }

        OpenXmlValidation.AssertValid(stream);

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);
        var footerPart = Assert.Single(package.MainDocumentPart!.FooterParts);

        Assert.Equal("Page footer", footerPart.Footer!.InnerText);
    }

    /// <summary>
    /// <c>CT_SectPr</c> puts the header and footer references before the page size and margins. The
    /// document is otherwise well-formed, so only schema validation catches the wrong order.
    /// </summary>
    [Fact]
    public void SectionProperties_KeepReferencesBeforePageSizeAndMargins()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            // Deliberately out of schema order: the page setup is applied before the header exists.
            document.ApplyPageSetup(new PageSetup { PaperSize = PaperSize.A4 }.WithUniformMargins(56));
            document.GetBody().AddParagraph("Body");
            document.AddHeader().AddParagraph("Header");
            document.AddFooter().AddParagraph("Footer");
        }

        OpenXmlValidation.AssertValid(stream);

        var children = ReadDocumentElement(
            stream,
            document => document.Body!.GetFirstChild<WordLib.SectionProperties>()!.ChildElements
                .Select(child => child.LocalName)
                .ToList());

        Assert.Equal(
            ["headerReference", "footerReference", "pgSz", "pgMar"],
            children);
    }

    /// <summary>
    /// A first-page header is valid markup without <c>w:titlePg</c> — it simply never appears, which is
    /// the kind of silent failure this library should absorb.
    /// </summary>
    [Fact]
    public void AddHeader_ForTheFirstPage_TurnsOnTheTitlePageSwitch()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            document.GetBody().AddParagraph("Body");
            document.AddHeader(HeaderFooterKind.First).AddParagraph("Letterhead");
        }

        OpenXmlValidation.AssertValid(stream);

        var sectionProperties = ReadDocumentElement(
            stream,
            document => document.Body!.GetFirstChild<WordLib.SectionProperties>()!);

        Assert.NotNull(sectionProperties.GetFirstChild<WordLib.TitlePage>());
        Assert.Equal(
            WordLib.HeaderFooterValues.First,
            sectionProperties.GetFirstChild<WordLib.HeaderReference>()!.Type?.Value);
    }

    /// <summary>
    /// The even-page equivalent lives in the document settings rather than the section, and is just as
    /// silently required.
    /// </summary>
    [Fact]
    public void AddHeader_ForEvenPages_TurnsOnTheEvenAndOddHeadersSetting()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            document.GetBody().AddParagraph("Body");
            document.AddHeader(HeaderFooterKind.Even).AddParagraph("Even page");
        }

        OpenXmlValidation.AssertValid(stream);

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);
        var settings = package.MainDocumentPart!.DocumentSettingsPart?.Settings;

        Assert.NotNull(settings);
        Assert.NotNull(settings.GetFirstChild<WordLib.EvenAndOddHeaders>());
    }

    [Fact]
    public void AddHeader_ForEachKind_CreatesOnePartPerKind()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            document.GetBody().AddParagraph("Body");
            document.AddHeader(HeaderFooterKind.Default).AddParagraph("default");
            document.AddHeader(HeaderFooterKind.First).AddParagraph("first");
            document.AddHeader(HeaderFooterKind.Even).AddParagraph("even");
        }

        OpenXmlValidation.AssertValid(stream);

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);

        Assert.Equal(3, package.MainDocumentPart!.HeaderParts.Count());
        Assert.Equal(3, package.MainDocumentPart.Document!.Descendants<WordLib.HeaderReference>().Count());
    }

    [Fact]
    public void AddHeader_CalledTwiceForOneKind_ReturnsTheSameHeader()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        document.GetBody().AddParagraph("Body");

        var first = document.AddHeader();
        var second = document.AddHeader();

        Assert.Same(first, second);
    }

    /// <summary>
    /// Reopening and asking for the header again has to find the existing part; creating a second one
    /// would leave the original content in the package but unreferenced.
    /// </summary>
    [Fact]
    public void AddHeader_AfterReopening_ReusesTheExistingHeader()
    {
        var filePath = GetFilepath("header-reopen.docx");

        using (var document = CreateDocument(filePath))
        {
            document.GetBody().AddParagraph("Body");
            document.AddHeader().AddParagraph("Original header");
        }

        using (var document = OpenDocument(filePath))
        {
            var header = document.AddHeader();

            Assert.Equal("Original header", header.GetAllTexts());
            header.AddParagraph("Second line");
        }

        OpenXmlValidation.AssertValid(filePath);

        using var package = WordprocessingDocument.Open(filePath, false);
        var headerPart = Assert.Single(package.MainDocumentPart!.HeaderParts);

        Assert.Single(package.MainDocumentPart.Document!.Descendants<WordLib.HeaderReference>());
        Assert.Contains("Second line", headerPart.Header!.InnerText, StringComparison.Ordinal);
    }

    /// <summary>
    /// A header is a block container, so everything the body accepts works there too — which is the
    /// point of sharing one implementation.
    /// </summary>
    [Fact]
    public void Header_AcceptsTablesAndImages()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            document.GetBody().AddParagraph("Body");

            var header = document.AddHeader();
            header.AddTable([["Left", "Right"]]);

            using var image = new MemoryStream(TestImages.MinimalPng());
            header.AddParagraph().AddImage(image);
        }

        OpenXmlValidation.AssertValid(stream);

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);
        var headerPart = package.MainDocumentPart!.HeaderParts.Single();

        Assert.Single(headerPart.Header!.Descendants<WordLib.Table>());
        Assert.Single(headerPart.ImageParts);
    }

    [Fact]
    public void HeadersAndFooters_ListsWhatWasAdded()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        document.GetBody().AddParagraph("Body");

        document.AddHeader();
        document.AddFooter(HeaderFooterKind.First);

        Assert.Equal(2, document.HeadersAndFooters.Count);
        Assert.Contains(document.HeadersAndFooters, item => item is { IsHeader: true, Kind: HeaderFooterKind.Default });
        Assert.Contains(document.HeadersAndFooters, item => item is { IsHeader: false, Kind: HeaderFooterKind.First });
    }

    [Fact]
    public void AddHeader_AfterClose_Throws()
    {
        using var stream = new MemoryStream();
        var document = CreateDocument(stream);
        document.GetBody().AddParagraph("Body");
        document.Close();

        Assert.Throws<ObjectDisposedException>(() => document.AddHeader());
    }
}
