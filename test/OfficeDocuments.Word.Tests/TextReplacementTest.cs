using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;
using OfficeDocuments.Word.TestKit;

namespace OfficeDocuments.Word.Tests;

/// <summary>
/// Text replacement, whose whole reason to exist is that the text to replace is usually not in one run.
/// </summary>
public class TextReplacementTest : WordTestBase
{
    /// <summary>
    /// The case the feature exists for. The document is built the way Word writes one — the placeholder
    /// split over three runs with a spell-check marker between them — so a per-run search finds nothing.
    /// </summary>
    [Fact]
    public void ReplaceText_PlaceholderSplitAcrossRuns_IsStillFound()
    {
        using var input = ForeignDocuments.WithSplitRuns(["Dear ", "{{customer", "}}, thank you."]);
        using var document = OpenDocument(input);
        var paragraph = document.GetBody().Paragraphs[0];

        Assert.Equal("Dear {{customer}}, thank you.", paragraph.GetTexts());

        var replaced = paragraph.ReplaceText("{{customer}}", "Ms Domanská");

        Assert.Equal(1, replaced);
        Assert.Equal("Dear Ms Domanská, thank you.", paragraph.GetTexts());
    }

    /// <summary>
    /// The replacement takes the formatting of the run where the match began, which is what makes a
    /// template fill look deliberate rather than like text pasted in from somewhere else.
    /// </summary>
    [Fact]
    public void ReplaceText_TakesTheFormattingOfTheRunTheMatchStartsIn()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            var paragraph = document.GetBody().AddParagraph();
            paragraph.AddText("Total: ");
            paragraph.AddText("{{amount", new TextFormat { Bold = true, Color = "C00000" });
            paragraph.AddText("}} CZK");

            Assert.Equal(1, paragraph.ReplaceText("{{amount}}", "1 240 000"));
        }

        OpenXmlValidation.AssertValid(stream);

        using var reopened = OpenDocument(stream, isEditable: false);
        var runs = reopened.GetBody().Paragraphs[0].Runs;

        Assert.Equal("Total: 1 240 000 CZK", reopened.GetBody().Paragraphs[0].GetTexts());
        Assert.Contains(runs, run => run.Text == "1 240 000" && run.Format.Bold == true && run.Format.Color == "C00000");
    }

    /// <summary>
    /// A run the replacement empties is removed, so repeatedly filling a template does not leave a
    /// growing trail of content-free runs behind.
    /// </summary>
    /// <remarks>
    /// Split so that the match starts in the first run and consumes the second one whole. The run the
    /// match starts in is the one that receives the replacement, so it is never the one left empty.
    /// </remarks>
    [Fact]
    public void ReplaceText_DropsTheRunsItEmpties()
    {
        using var input = ForeignDocuments.WithSplitRuns(["A {{to", "ken}}", " B"]);
        using var document = OpenDocument(input);
        var paragraph = document.GetBody().Paragraphs[0];

        Assert.Equal(3, paragraph.Runs.Count);
        Assert.Equal(1, paragraph.ReplaceText("{{token}}", "value"));

        Assert.Equal(2, paragraph.Runs.Count);
        Assert.Equal("A value B", paragraph.GetTexts());
    }

    [Fact]
    public void ReplaceText_SeveralOccurrencesInOneRun_ReplacesAllOfThem()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var paragraph = document.GetBody().AddParagraph("one X two X three X four");

        Assert.Equal(3, paragraph.ReplaceText("X", "-"));
        Assert.Equal("one - two - three - four", paragraph.GetTexts());
    }

    /// <summary>
    /// Replacing occurrences right to left is what keeps this correct: the earlier offsets stay valid
    /// because every edit changes only the text after its own start.
    /// </summary>
    [Fact]
    public void ReplaceText_WithLongerText_KeepsTheLaterOccurrencesIntact()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var paragraph = document.GetBody().AddParagraph("a|b|c");

        Assert.Equal(2, paragraph.ReplaceText("|", " and "));
        Assert.Equal("a and b and c", paragraph.GetTexts());
    }

    [Fact]
    public void ReplaceText_WithEmptyText_DeletesTheMatch()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var paragraph = document.GetBody().AddParagraph("Draft — internal only");

        Assert.Equal(1, paragraph.ReplaceText(" — internal only", string.Empty));
        Assert.Equal("Draft", paragraph.GetTexts());
    }

    /// <summary>
    /// A newline in the replacement becomes <c>w:br</c>, the same as it does when text is authored, so a
    /// caller cannot tell from the markup which path produced it.
    /// </summary>
    [Fact]
    public void ReplaceText_WithNewline_ProducesALineBreak()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            var paragraph = document.GetBody().AddParagraph("Line A / Line B");
            Assert.Equal(1, paragraph.ReplaceText(" / ", "\n"));
        }

        OpenXmlValidation.AssertValid(stream);

        var breakCount = ReadDocumentElement(stream, document =>
            document.Descendants<DocumentFormat.OpenXml.Wordprocessing.Break>().Count());

        Assert.Equal(1, breakCount);

        using var reopened = OpenDocument(stream, isEditable: false);
        Assert.Equal("Line A\nLine B", reopened.GetBody().Paragraphs[0].GetTexts());
    }

    /// <summary>
    /// A line break reads as <c>\n</c>, so a match may run through one. The break element is then part of
    /// what gets replaced.
    /// </summary>
    [Fact]
    public void ReplaceText_MatchSpanningALineBreak_ReplacesTheBreakToo()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var paragraph = document.GetBody().AddParagraph("first\nsecond");

        Assert.Equal(1, paragraph.ReplaceText("first\nsecond", "joined"));
        Assert.Equal("joined", paragraph.GetTexts());
    }

    [Fact]
    public void ReplaceText_IgnoringCase_MatchesRegardlessOfCase()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var paragraph = document.GetBody().AddParagraph("Word and WORD and word");

        Assert.Equal(3, paragraph.ReplaceText("word", "term", StringComparison.OrdinalIgnoreCase));
        Assert.Equal("term and term and term", paragraph.GetTexts());
    }

    /// <summary>
    /// Trailing whitespace only survives an XML round trip with <c>xml:space="preserve"</c>, and text
    /// arriving through a replacement needs it just as much as authored text does.
    /// </summary>
    [Fact]
    public void ReplaceText_LeavingTrailingWhitespace_KeepsItSignificant()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            var paragraph = document.GetBody().AddParagraph("Total:#");
            Assert.Equal(1, paragraph.ReplaceText("#", " "));
        }

        Assert.Contains("xml:space=\"preserve\"", ReadMainDocumentXml(stream), StringComparison.Ordinal);

        using var reopened = OpenDocument(stream, isEditable: false);
        Assert.Equal("Total: ", reopened.GetBody().Paragraphs[0].GetTexts());
    }

    /// <summary>
    /// Two paragraphs are two texts. A phrase that only exists by reading across the boundary between
    /// them is not a match, because replacing it would have to merge or delete a paragraph.
    /// </summary>
    [Fact]
    public void ReplaceText_AcrossAParagraphBoundary_DoesNotMatch()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var body = document.GetBody();
        body.AddParagraph("end of one");
        body.AddParagraph("start of two");

        Assert.Equal(0, body.ReplaceText("one start", "merged"));
        Assert.Equal("end of one\nstart of two", body.GetAllTexts());
    }

    [Fact]
    public void ReplaceText_TextInsideAHyperlink_IsReplaced()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            var body = document.GetBody();
            body.AddParagraph().AddHyperlink("{{portal}}", "https://example.com");

            Assert.Equal(1, body.ReplaceText("{{portal}}", "the portal"));
        }

        OpenXmlValidation.AssertValid(stream);

        using var reopened = OpenDocument(stream, isEditable: false);
        Assert.Equal("the portal", reopened.GetBody().GetAllTexts());
    }

    [Fact]
    public void ReplaceText_OnAContainer_ReachesTableCellsAndNestedTables()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            var body = document.GetBody();
            body.AddParagraph("Report for {{month}}");

            var table = body.AddTable([["Period", "{{month}}"]]);
            table.GetCell(0, 0).AddTable([["nested {{month}}"]]);

            Assert.Equal(3, body.ReplaceText("{{month}}", "July"));
        }

        OpenXmlValidation.AssertValid(stream);

        using var reopened = OpenDocument(stream, isEditable: false);
        var text = reopened.GetBody().GetAllTexts();

        Assert.DoesNotContain("{{month}}", text, StringComparison.Ordinal);
        Assert.Contains("Report for July", text, StringComparison.Ordinal);
        Assert.Contains("nested July", text, StringComparison.Ordinal);
    }

    /// <summary>
    /// The document-level overload is the one a template fill wants: a running header holding a date or a
    /// customer name is exactly what a body-only pass leaves behind.
    /// </summary>
    [Fact]
    public void ReplaceText_OnTheDocument_ReachesHeadersAndFooters()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            document.AddHeader().AddParagraph("{{client}} — confidential");
            document.AddFooter().AddParagraph("Prepared for {{client}}");
            document.GetBody().AddParagraph("Dear {{client}},");

            Assert.Equal(3, document.ReplaceText("{{client}}", "Acme"));
        }

        OpenXmlValidation.AssertValid(stream);

        using var reopened = OpenDocument(stream, isEditable: false);

        Assert.Equal("Dear Acme,", reopened.GetBody().GetAllTexts());
        Assert.All(reopened.HeadersAndFooters,
            container => Assert.DoesNotContain("{{client}}", container.GetAllTexts(), StringComparison.Ordinal));
        Assert.Contains(reopened.HeadersAndFooters,
            container => container.GetAllTexts() == "Acme — confidential");
    }

    [Fact]
    public void ReplaceText_WithNoMatch_ReportsNothingAndLeavesTheTextAlone()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var body = document.GetBody();
        body.AddParagraph("nothing to see here");

        Assert.Equal(0, body.ReplaceText("{{missing}}", "x"));
        Assert.Equal("nothing to see here", body.GetAllTexts());
    }

    [Fact]
    public void ReplaceText_WithEmptySearchText_Throws()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var paragraph = document.GetBody().AddParagraph("text");

        Assert.Throws<ArgumentException>(() => paragraph.ReplaceText(string.Empty, "x"));
    }

    /// <summary>
    /// Replacing text repeatedly has to stay stable: the second pass runs over markup the first one
    /// restructured, which is where an implementation that mutates while enumerating comes apart.
    /// </summary>
    [Fact]
    public void ReplaceText_AppliedRepeatedly_StaysCorrect()
    {
        using var input = ForeignDocuments.WithSplitRuns(["{{a", "}} and ", "{{b", "}}"]);
        using var document = OpenDocument(input);
        var body = document.GetBody();

        Assert.Equal(1, body.ReplaceText("{{a}}", "first"));
        Assert.Equal(1, body.ReplaceText("{{b}}", "second"));
        Assert.Equal(1, body.ReplaceText("first and second", "both"));
        Assert.Equal("both", body.GetAllTexts());
    }

    /// <summary>
    /// Setting a paragraph's text clears its content but not its formatting, which is the difference
    /// between filling a styled placeholder line and flattening it.
    /// </summary>
    [Fact]
    public void SetText_ReplacesContentAndKeepsParagraphFormatting()
    {
        using var stream = new MemoryStream();

        using (var document = CreateDocument(stream))
        {
            var paragraph = document.GetBody()
                .AddParagraph("placeholder", new ParagraphFormat { StyleId = WordStyleIds.Heading1, Alignment = ParagraphAlignment.Center });

            paragraph.AddHyperlink("link", "https://example.com");
            paragraph.SetText("Actual heading");
        }

        OpenXmlValidation.AssertValid(stream);

        using var reopened = OpenDocument(stream, isEditable: false);
        var read = reopened.GetBody().Paragraphs[0];

        Assert.Equal("Actual heading", read.GetTexts());
        Assert.Equal(WordStyleIds.Heading1, read.Format.StyleId);
        Assert.Equal(ParagraphAlignment.Center, read.Format.Alignment);
    }
}
