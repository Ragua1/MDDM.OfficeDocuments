using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.Tests;

/// <summary>
/// Covers hyperlinks: the relationship, the markup, and the style that makes a link look like one.
/// </summary>
public class HyperlinkTest : WordTestBase
{
    [Fact]
    public void AddHyperlink_CreatesAnExternalRelationshipToTheTarget()
    {
        using var stream = WriteAndValidate(body =>
            body.AddParagraph().AddHyperlink("Example", "https://example.com/docs"));

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);
        var relationship = Assert.Single(package.MainDocumentPart!.HyperlinkRelationships);

        Assert.Equal("https://example.com/docs", relationship.Uri.ToString());
        Assert.True(relationship.IsExternal);
    }

    /// <summary>
    /// The <c>w:hyperlink</c> element points at the relationship by id; a mismatch produces a link
    /// that goes nowhere.
    /// </summary>
    [Fact]
    public void AddHyperlink_ReferencesTheRelationshipItCreated()
    {
        using var stream = WriteAndValidate(body =>
            body.AddParagraph().AddHyperlink("Example", "https://example.com"));

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);
        var relationshipId = package.MainDocumentPart!.HyperlinkRelationships.Single().Id;
        var hyperlink = package.MainDocumentPart.Document!.Descendants<WordLib.Hyperlink>().Single();

        Assert.Equal(relationshipId, hyperlink.Id?.Value);
    }

    [Fact]
    public void AddHyperlink_KeepsDisplayTextSeparateFromTheTarget()
    {
        using var stream = WriteAndValidate(body =>
            body.AddParagraph().AddHyperlink("our documentation", "https://example.com/very/long/path"));

        using var document = OpenDocument(stream, isEditable: false);

        Assert.Equal("our documentation", document.GetBody().GetAllTexts());
    }

    /// <summary>
    /// Referencing the Hyperlink style is not enough; without a definition the link renders as plain
    /// black text.
    /// </summary>
    [Fact]
    public void AddHyperlink_DefinesTheHyperlinkCharacterStyle()
    {
        using var stream = WriteAndValidate(body =>
            body.AddParagraph().AddHyperlink("Example", "https://example.com"));

        var styles = ReadStylesXml(stream);

        Assert.Contains("w:styleId=\"Hyperlink\"", styles, StringComparison.Ordinal);
        Assert.Contains("w:type=\"character\"", styles, StringComparison.Ordinal);
    }

    /// <summary>
    /// A hyperlink wraps its run in a container, so a paragraph that read only its direct children
    /// would report no runs and no text for a link.
    /// </summary>
    [Fact]
    public void Runs_IncludeRunsNestedInsideAHyperlink()
    {
        using var stream = WriteAndValidate(body => body
            .AddParagraph()
            .AddText("See ")
            .AddHyperlink("the docs", "https://example.com")
            .AddText(" for details."));

        using var document = OpenDocument(stream, isEditable: false);
        var paragraph = document.GetBody().Paragraphs[0];

        Assert.Equal(3, paragraph.Runs.Count);
        Assert.Equal("the docs", paragraph.Runs[1].Text);
        Assert.Equal("See the docs for details.", paragraph.GetTexts());
    }

    [Fact]
    public void AddHyperlink_AppliesTheHyperlinkStyleToItsRun()
    {
        using var stream = WriteAndValidate(body =>
            body.AddParagraph().AddHyperlink("Example", "https://example.com"));

        using var document = OpenDocument(stream, isEditable: false);

        Assert.Equal(WordStyleIds.Hyperlink, document.GetBody().Paragraphs[0].Runs[0].Format.StyleId);
    }

    /// <summary>
    /// The caller's format layers on top of the hyperlink style rather than replacing it, so a
    /// recoloured link is still underlined and still a link.
    /// </summary>
    [Fact]
    public void AddHyperlink_WithFormat_LayersItOverTheHyperlinkStyle()
    {
        using var stream = WriteAndValidate(body => body
            .AddParagraph()
            .AddHyperlink("Example", "https://example.com", new TextFormat { Bold = true, Color = "FF0000" }));

        using var document = OpenDocument(stream, isEditable: false);
        var format = document.GetBody().Paragraphs[0].Runs[0].Format;

        Assert.Equal(WordStyleIds.Hyperlink, format.StyleId);
        Assert.True(format.Bold);
        Assert.Equal("FF0000", format.Color);
    }

    [Fact]
    public void AddHyperlink_ToMailtoAddress_IsAccepted()
    {
        using var stream = WriteAndValidate(body =>
            body.AddParagraph().AddHyperlink("Write to us", "mailto:info@example.com"));

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);

        Assert.Equal("mailto:info@example.com", package.MainDocumentPart!.HyperlinkRelationships.Single().Uri.ToString());
    }

    [Fact]
    public void AddHyperlink_InATableCell_IsSchemaValid()
    {
        using var stream = WriteAndValidate(body =>
        {
            var cell = body.AddTable(1, 1).GetCell(0, 0);
            cell.Paragraphs[0].AddHyperlink("Example", "https://example.com");
        });

        using var document = OpenDocument(stream, isEditable: false);

        Assert.Equal("Example", document.GetBody().Tables[0].GetAllTexts());
    }

    [Fact]
    public void AddHyperlink_MultipleLinks_EachGetsItsOwnRelationship()
    {
        using var stream = WriteAndValidate(body =>
        {
            body.AddParagraph().AddHyperlink("First", "https://example.com/one");
            body.AddParagraph().AddHyperlink("Second", "https://example.com/two");
        });

        stream.Position = 0;
        using var package = WordprocessingDocument.Open(stream, false);
        var relationships = package.MainDocumentPart!.HyperlinkRelationships.ToList();

        Assert.Equal(2, relationships.Count);
        Assert.Equal(2, relationships.Select(relationship => relationship.Id).Distinct().Count());
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("not a uri")]
    [InlineData("/relative/path")]
    public void AddHyperlink_WithInvalidUrl_Throws(string url)
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var paragraph = document.GetBody().AddParagraph();

        Assert.Throws<ArgumentException>(() => paragraph.AddHyperlink("text", url));
    }

    [Fact]
    public void AddHyperlink_WithNullText_Throws()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var paragraph = document.GetBody().AddParagraph();

        Assert.Throws<ArgumentNullException>(() => paragraph.AddHyperlink(null!, "https://example.com"));
    }

    /// <summary>
    /// Character styles other than the built-in ones are the caller's to define, so an unknown one is
    /// written through rather than invented.
    /// </summary>
    [Fact]
    public void TextFormatStyleId_WithUnknownStyle_IsWrittenWithoutBeingDefined()
    {
        using var stream = WriteAndValidate(body =>
            body.AddParagraph().AddText("x", new TextFormat { StyleId = "CompanyEmphasis" }));

        var run = ReadDocumentElement(stream, document => document.Descendants<WordLib.RunStyle>().Single());

        Assert.Equal("CompanyEmphasis", run.Val?.Value);
        Assert.DoesNotContain("w:styleId=\"CompanyEmphasis\"", ReadStylesXml(stream), StringComparison.Ordinal);
    }
}
