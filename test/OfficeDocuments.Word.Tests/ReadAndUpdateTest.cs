using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;
using OfficeDocuments.Word.TestKit;

namespace OfficeDocuments.Word.Tests;

/// <summary>
/// Opening a document that already exists, changing it, and saving it back.
/// </summary>
/// <remarks>
/// The authoring tests all start from an empty package, which is the easy half. These start from a
/// document with content, section properties, headers, and relationships already in it — the state
/// every template workflow actually begins in.
/// </remarks>
public class ReadAndUpdateTest : WordTestBase
{
    /// <summary>
    /// The scenario the task exists for: a template with placeholders in the body, in a table, and in a
    /// running header, filled in one call and saved back as a valid document.
    /// </summary>
    [Fact]
    public void FillTemplate_ReplacesEveryPlaceholderAndStaysValid()
    {
        var filePath = GetFilepath("template.docx");

        using (var template = CreateDocument(filePath))
        {
            template.AddHeader().AddParagraph("{{company}} — {{month}}");
            template.AddFooter().AddParagraph("Issued {{date}}");

            var body = template.GetBody();
            body.AddParagraph("Statement for {{company}}", new ParagraphFormat { StyleId = WordStyleIds.Title });
            body.AddTable([
                ["Period", "{{month}}"],
                ["Issued", "{{date}}"],
            ]);
            body.AddParagraph("Questions? Write to {{contact}}.");
        }

        OpenXmlValidation.AssertValid(filePath);

        using (var filled = OpenDocument(filePath))
        {
            Assert.Equal(2, filled.ReplaceText("{{company}}", "Acme s.r.o."));
            Assert.Equal(2, filled.ReplaceText("{{month}}", "July 2026"));
            Assert.Equal(2, filled.ReplaceText("{{date}}", "2026-07-27"));
            Assert.Equal(1, filled.ReplaceText("{{contact}}", "support@example.com"));
        }

        OpenXmlValidation.AssertValid(filePath);

        using var result = OpenDocument(filePath, isEditable: false);
        var text = result.GetBody().GetAllTexts();

        Assert.DoesNotContain("{{", text, StringComparison.Ordinal);
        Assert.Contains("Statement for Acme s.r.o.", text, StringComparison.Ordinal);
        Assert.Contains("Period\tJuly 2026", text, StringComparison.Ordinal);
        Assert.All(result.HeadersAndFooters,
            container => Assert.DoesNotContain("{{", container.GetAllTexts(), StringComparison.Ordinal));
    }

    /// <summary>
    /// A document opened from disk has to report the headers and footers it already contains. Reporting
    /// only the ones added in this session made every read and template workflow miss them, and nothing
    /// about the resulting document looked wrong — the headers were simply never visited.
    /// </summary>
    [Fact]
    public void HeadersAndFooters_OnAnOpenedDocument_ReportsTheOnesTheDocumentAlreadyHad()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            document.AddHeader().AddParagraph("default header");
            document.AddHeader(HeaderFooterKind.First).AddParagraph("first header");
            document.AddFooter().AddParagraph("default footer");
        }

        using var reopened = OpenDocument(stream, isEditable: false);
        var containers = reopened.HeadersAndFooters;

        Assert.Equal(3, containers.Count);
        Assert.Equal(2, containers.Count(container => container.IsHeader));
        Assert.Contains(containers, container => container is { IsHeader: true, Kind: HeaderFooterKind.First });
        Assert.Contains(containers, container => container.GetAllTexts() == "default footer");
    }

    /// <summary>
    /// Asking for a header a document already has returns that header with its content, rather than
    /// adding a second reference and orphaning the first.
    /// </summary>
    [Fact]
    public void AddHeader_OnAnOpenedDocumentThatHasOne_ReturnsTheExistingHeader()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            document.AddHeader().AddParagraph("original");
        }

        using (var document = OpenDocument(stream))
        {
            var header = document.AddHeader();

            Assert.Equal("original", header.GetAllTexts());

            header.AddParagraph("added later");

            Assert.Single(document.HeadersAndFooters);
        }

        OpenXmlValidation.AssertValid(stream);

        using var reopened = OpenDocument(stream, isEditable: false);

        Assert.Single(reopened.HeadersAndFooters);
        Assert.Equal("original\nadded later", reopened.HeadersAndFooters[0].GetAllTexts());
    }

    /// <summary>
    /// The same wrapper instance comes back whether it was created by <c>AddHeader</c> or discovered by
    /// reading the document, so a caller cannot end up holding two facades over one header.
    /// </summary>
    [Fact]
    public void HeadersAndFooters_ReturnsTheSameInstanceAsAddHeader()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var added = document.AddHeader();

        Assert.Same(added, document.HeadersAndFooters.Single());
    }

    /// <summary>
    /// Editing an existing document keeps everything it already had: its styles, its numbering, its
    /// relationships, and its page setup.
    /// </summary>
    [Fact]
    public void EditExistingDocument_KeepsWhatItAlreadyContained()
    {
        var filePath = GetFilepath("existing.docx");

        using (var document = CreateDocument(filePath))
        {
            document.ApplyPageSetup(new PageSetup { PaperSize = PaperSize.A5 }.WithUniformMargins(40));
            document.SetMetadata(new DocumentMetadata { Title = "Original", Author = "First" });

            var body = document.GetBody();
            body.AddHeading("Kept heading", 1);
            body.AddListItem("kept item", ListStyle.Number);
            body.AddParagraph().AddHyperlink("kept link", "https://example.com/kept");
        }

        using (var document = OpenDocument(filePath))
        {
            var body = document.GetBody();

            Assert.Equal(PaperSize.A5, document.PageSetup.PaperSize);
            Assert.Equal("Original", document.Metadata.Title);
            Assert.Equal(WordStyleIds.Heading1, body.Paragraphs[0].Format.StyleId);
            Assert.Equal(ListStyle.Number, body.Paragraphs[1].Format.ListStyle);

            body.AddListItem("added item", ListStyle.Number);
            body.AddParagraph().AddHyperlink("added link", "https://example.com/added");
            document.SetMetadata(new DocumentMetadata { LastModifiedBy = "Second" });
        }

        OpenXmlValidation.AssertValid(filePath);

        using var reopened = OpenDocument(filePath, isEditable: false);
        var reopenedBody = reopened.GetBody();

        Assert.Equal("Original", reopened.Metadata.Title);
        Assert.Equal("Second", reopened.Metadata.LastModifiedBy);
        Assert.Equal(PaperSize.A5, reopened.PageSetup.PaperSize);

        // Both list items share one numbering definition: the existing one is reused rather than
        // duplicated, which is what keeps an appended-to document from growing a definition per edit.
        var listItems = reopenedBody.Paragraphs.Where(paragraph => paragraph.Format.ListStyle == ListStyle.Number).ToList();

        Assert.Equal(2, listItems.Count);
        Assert.Contains("added link", reopenedBody.GetAllTexts(), StringComparison.Ordinal);
        Assert.Contains("kept link", reopenedBody.GetAllTexts(), StringComparison.Ordinal);
    }

    /// <summary>
    /// Removing a placeholder block is the other half of a template workflow: an optional section that
    /// this run does not need has to be able to go away entirely.
    /// </summary>
    [Fact]
    public void RemoveOptionalSection_LeavesTheRestOfTheDocumentIntact()
    {
        var filePath = GetFilepath("optional.docx");

        using (var document = CreateDocument(filePath))
        {
            var body = document.GetBody();
            body.AddHeading("Always", 1);
            body.AddParagraph("Always present.");
            body.AddHeading("Optional", 1);
            body.AddTable([["only", "if needed"]]);
        }

        using (var document = OpenDocument(filePath))
        {
            var body = document.GetBody();
            var optionalHeading = body.FindParagraphs("Optional").Single();

            Assert.True(body.Remove(optionalHeading));
            Assert.True(body.Remove(body.Tables[0]));
        }

        OpenXmlValidation.AssertValid(filePath);

        using var reopened = OpenDocument(filePath, isEditable: false);

        Assert.Equal("Always\nAlways present.", reopened.GetBody().GetAllTexts());
        Assert.Empty(reopened.GetBody().Tables);
    }

    /// <summary>
    /// Opening a document read-only must leave the file untouched, byte for byte. A save on close would
    /// rewrite the package even when nothing changed, which is a surprising thing for a read to do to a
    /// file under version control or on a shared drive.
    /// </summary>
    [Fact]
    public void OpenReadOnly_DoesNotRewriteTheFile()
    {
        var filePath = GetFilepath("readonly.docx");

        using (var document = CreateDocument(filePath))
        {
            document.GetBody().AddParagraph("untouched");
        }

        var before = File.ReadAllBytes(filePath);

        using (var document = OpenDocument(filePath, isEditable: false))
        {
            Assert.Equal("untouched", document.GetBody().GetAllTexts());
        }

        Assert.Equal(before, File.ReadAllBytes(filePath));
    }

    /// <summary>
    /// A document written by another producer is the real input case. Its runs are split where Word chose
    /// to split them, and appending, replacing, and reading all have to work over that.
    /// </summary>
    [Fact]
    public void UpdateForeignDocument_ReadsAppendsAndReplacesOverSplitRuns()
    {
        using var stream = ForeignDocuments.WithSplitRuns(
            ["Contract with ", "{{party", "}} signed on {{date}}."],
            ["Prepared by ", "{{author", "}}."]);

        using (var document = OpenDocument(stream))
        {
            var body = document.GetBody();

            Assert.Equal(2, body.Paragraphs.Count);
            Assert.Equal(3, body.Paragraphs[0].Runs.Count);

            Assert.Equal(1, body.ReplaceText("{{party}}", "Acme s.r.o."));
            Assert.Equal(1, body.ReplaceText("{{date}}", "2026-07-27"));
            Assert.Equal(1, body.ReplaceText("{{author}}", "M. Domanský"));

            body.AddParagraph("Appended after the section properties existed.");
        }

        OpenXmlValidation.AssertValid(stream);

        using var reopened = OpenDocument(stream, isEditable: false);

        Assert.Equal(
            "Contract with Acme s.r.o. signed on 2026-07-27.\nPrepared by M. Domanský.\nAppended after the section properties existed.",
            reopened.GetBody().GetAllTexts());
    }

    /// <summary>
    /// Discarding changes has to actually discard them, including on a document opened from disk.
    /// </summary>
    [Fact]
    public void CloseWithoutSaving_LeavesTheDocumentAsItWas()
    {
        var filePath = GetFilepath("discarded.docx");

        using (var document = CreateDocument(filePath))
        {
            document.GetBody().AddParagraph("original");
        }

        var document2 = OpenDocument(filePath);
        document2.GetBody().AddParagraph("should not survive");
        document2.Close(saveDocument: false);

        using var reopened = OpenDocument(filePath, isEditable: false);

        Assert.Equal("original", reopened.GetBody().GetAllTexts());
    }
}
