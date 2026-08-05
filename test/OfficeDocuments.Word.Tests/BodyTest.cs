using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Word.Enums;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.Tests;

/// <summary>
/// Covers the body's block model and the text it reads back.
/// </summary>
public class BodyTest : WordTestBase
{
    /// <summary>
    /// Regression test: the paragraph list used to be a snapshot taken when the body was wrapped, so
    /// paragraphs added afterwards were invisible and <c>GetAllTexts</c> returned an empty string.
    /// </summary>
    [Fact]
    public void AddParagraph_IsVisibleImmediately()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var body = document.GetBody();

        body.AddParagraph().AddText("Hello world");

        Assert.Single(body.Paragraphs);
        Assert.Equal("Hello world", body.GetAllTexts());
    }

    [Fact]
    public void Paragraphs_ReturnsSameInstanceForTheSameParagraph()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var body = document.GetBody();
        var added = body.AddParagraph("First");

        Assert.Same(added, body.Paragraphs[0]);
    }

    [Fact]
    public void Paragraphs_AfterReopen_AreInDocumentOrder()
    {
        using var stream = WriteAndValidate(body =>
        {
            body.AddParagraph("One");
            body.AddParagraph("Two");
            body.AddParagraph("Three");
        });

        using var document = OpenDocument(stream, isEditable: false);
        var paragraphs = document.GetBody().Paragraphs;

        Assert.Equal(["One", "Two", "Three"], paragraphs.Select(paragraph => paragraph.GetTexts()));
    }

    /// <summary>
    /// An empty paragraph is a blank line in the document, so dropping it while reading loses
    /// structure that the author put there deliberately.
    /// </summary>
    [Fact]
    public void GetAllTexts_KeepsEmptyParagraphs()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var body = document.GetBody();

        body.AddParagraph("Above");
        body.AddParagraph();
        body.AddParagraph("Below");

        Assert.Equal("Above\n\nBelow", body.GetAllTexts());
    }

    /// <summary>
    /// Regression test: text used to be trimmed on the way out, so adjacent runs ran together.
    /// </summary>
    [Fact]
    public void AddText_WithSurroundingSpaces_RoundTripsExactly()
    {
        using var stream = WriteAndValidate(body =>
        {
            var paragraph = body.AddParagraph();
            paragraph.AddText("Total: ");
            paragraph.AddText("42 ");
            paragraph.AddText("CZK");
        });

        using var document = OpenDocument(stream, isEditable: false);

        Assert.Equal("Total: 42 CZK", document.GetBody().GetAllTexts());
        Assert.Contains("xml:space=\"preserve\"", ReadMainDocumentXml(stream), StringComparison.Ordinal);
    }

    [Fact]
    public void AddText_WithNewlines_WritesLineBreaks()
    {
        using var stream = WriteAndValidate(body => body.AddParagraph("first\nsecond"));

        Assert.Contains("<w:br />", ReadMainDocumentXml(stream), StringComparison.Ordinal);

        using var document = OpenDocument(stream, isEditable: false);
        Assert.Equal("first\nsecond", document.GetBody().GetAllTexts());
    }

    [Fact]
    public void AddBreak_PageBreak_IsReadBackAsLineSeparator()
    {
        using var stream = WriteAndValidate(body => body
            .AddParagraph()
            .AddText("before")
            .AddBreak(BreakType.Page)
            .AddText("after"));

        using var document = OpenDocument(stream, isEditable: false);

        Assert.Equal("before\nafter", document.GetBody().GetAllTexts());
    }

    [Fact]
    public void GetTextElements_ReturnsOneElementPerTextRun()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var paragraph = document.GetBody().AddParagraph();
        paragraph.AddText("alpha");
        paragraph.AddText("beta");

        Assert.Equal(["alpha", "beta"], paragraph.GetTextElements().Select(element => element.TextValue));
    }

    [Fact]
    public void AddParagraph_WithNullText_Throws()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var body = document.GetBody();

        Assert.Throws<ArgumentNullException>(() => body.AddParagraph((string)null!));
    }

    /// <summary>
    /// Every real document carries a trailing <c>w:sectPr</c>, which the schema requires to stay the
    /// last child of <c>w:body</c>. Appending a paragraph past it produces a file Word has to repair.
    /// </summary>
    [Fact]
    public void AddParagraph_OnDocumentWithSectionProperties_KeepsThemLast()
    {
        using var stream = new MemoryStream();
        CreateDocumentWithSectionProperties(stream);

        using (var document = OpenDocument(stream))
        {
            document.GetBody().AddParagraph("Appended after opening");
        }

        OpenXmlValidation.AssertValid(stream);

        var body = ReadDocumentElement(stream, document => document.Body!);

        Assert.IsType<WordLib.SectionProperties>(body.LastChild);
        Assert.Equal(2, body.Elements<WordLib.Paragraph>().Count());
    }

    /// <summary>
    /// Builds a document the way Word does, rather than the way this library does, so the test starts
    /// from markup the library did not produce.
    /// </summary>
    private static void CreateDocumentWithSectionProperties(Stream stream)
    {
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var body = new WordLib.Body();
        body.AppendChild(new WordLib.Paragraph(new WordLib.Run(new WordLib.Text("Existing content"))));
        body.AppendChild(new WordLib.SectionProperties());

        document.AddMainDocumentPart().Document = new WordLib.Document(body);
    }
}
