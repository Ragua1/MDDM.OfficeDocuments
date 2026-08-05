using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.Tests;

/// <summary>
/// Covers paragraph-level formatting and the built-in styles the library defines on demand.
/// </summary>
public class ParagraphFormattingTest : WordTestBase
{
    [Fact]
    public void ApplyFormat_RoundTripsEveryModelledProperty()
    {
        var format = new ParagraphFormat
        {
            Alignment = ParagraphAlignment.Center,
            SpacingBefore = 6,
            SpacingAfter = 12,
            LineSpacing = 1.5,
            IndentLeft = 36,
            IndentRight = 18,
            IndentFirstLine = 24,
        };

        using var stream = WriteAndValidate(body => body.AddParagraph("centered", format));

        using var document = OpenDocument(stream, isEditable: false);
        var actual = document.GetBody().Paragraphs[0].Format;

        Assert.Equal(ParagraphAlignment.Center, actual.Alignment);
        Assert.Equal(6, actual.SpacingBefore);
        Assert.Equal(12, actual.SpacingAfter);
        Assert.Equal(1.5, actual.LineSpacing);
        Assert.Equal(36, actual.IndentLeft);
        Assert.Equal(18, actual.IndentRight);
        Assert.Equal(24, actual.IndentFirstLine);
    }

    /// <summary>
    /// Points are the library's unit; twentieths of a point are the document's. 18 pt is 360 twips.
    /// </summary>
    [Fact]
    public void Spacing_IsWrittenInTwips()
    {
        using var stream = WriteAndValidate(body => body.AddParagraph("x", new ParagraphFormat { SpacingBefore = 18 }));

        Assert.Contains("w:before=\"360\"", ReadMainDocumentXml(stream), StringComparison.Ordinal);
    }

    /// <summary>
    /// Explicit spacing has to switch auto-spacing off, or a style that turns it on wins and the
    /// requested spacing is silently ignored by Word.
    /// </summary>
    [Fact]
    public void Spacing_DisablesAutoSpacing()
    {
        using var stream = WriteAndValidate(body => body.AddParagraph("x", new ParagraphFormat { SpacingBefore = 6, SpacingAfter = 6 }));

        var spacing = ReadDocumentElement(stream, document => document.Descendants<WordLib.SpacingBetweenLines>().Single());

        Assert.NotNull(spacing.BeforeAutoSpacing);
        Assert.NotNull(spacing.AfterAutoSpacing);
        Assert.False(spacing.BeforeAutoSpacing.Value);
        Assert.False(spacing.AfterAutoSpacing.Value);
    }

    /// <summary>
    /// A negative first-line indent is a hanging indent, which the document stores as a different
    /// attribute rather than as a negative number.
    /// </summary>
    [Fact]
    public void IndentFirstLine_Negative_BecomesHangingIndent()
    {
        using var stream = WriteAndValidate(body => body.AddParagraph("x", new ParagraphFormat { IndentFirstLine = -18 }));
        var xml = ReadMainDocumentXml(stream);

        Assert.Contains("w:hanging=\"360\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("w:firstLine=", xml, StringComparison.Ordinal);

        using var document = OpenDocument(stream, isEditable: false);
        Assert.Equal(-18, document.GetBody().Paragraphs[0].Format.IndentFirstLine);
    }

    [Fact]
    public void ApplyFormat_Twice_KeepsPropertiesItDoesNotSet()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var paragraph = document.GetBody().AddParagraph("x", new ParagraphFormat { Alignment = ParagraphAlignment.Right });

        paragraph.ApplyFormat(new ParagraphFormat { SpacingAfter = 10 });

        Assert.Equal(ParagraphAlignment.Right, paragraph.Format.Alignment);
        Assert.Equal(10, paragraph.Format.SpacingAfter);
    }

    [Theory]
    [InlineData(ParagraphAlignment.Left)]
    [InlineData(ParagraphAlignment.Center)]
    [InlineData(ParagraphAlignment.Right)]
    [InlineData(ParagraphAlignment.Justify)]
    public void Alignment_RoundTrips(ParagraphAlignment alignment)
    {
        using var stream = WriteAndValidate(body => body.AddParagraph("x", new ParagraphFormat { Alignment = alignment }));

        using var document = OpenDocument(stream, isEditable: false);

        Assert.Equal(alignment, document.GetBody().Paragraphs[0].Format.Alignment);
    }

    /// <summary>
    /// Referencing a style is not enough: without a definition in the styles part, a heading renders
    /// as body text. The library has to add the definition for the reference to mean anything.
    /// </summary>
    [Fact]
    public void AddHeading_DefinesTheStyleItReferences()
    {
        using var stream = WriteAndValidate(body => body.AddHeading("Chapter one", 1));

        Assert.Contains("w:val=\"Heading1\"", ReadMainDocumentXml(stream), StringComparison.Ordinal);

        var styles = ReadStylesXml(stream);
        Assert.Contains("w:styleId=\"Heading1\"", styles, StringComparison.Ordinal);
        Assert.Contains("w:val=\"heading 1\"", styles, StringComparison.Ordinal);
    }

    /// <summary>
    /// The outline level is what puts a heading in Word's navigation pane and in a generated table of
    /// contents, so it is part of the feature rather than decoration.
    /// </summary>
    [Fact]
    public void AddHeading_SetsTheOutlineLevel()
    {
        using var stream = WriteAndValidate(body => body.AddHeading("Chapter two", 2));

        Assert.Contains("<w:outlineLvl w:val=\"1\" />", ReadStylesXml(stream), StringComparison.Ordinal);
    }

    /// <summary>
    /// Every built-in style is based on Normal, so Normal has to exist too or the chain dangles.
    /// </summary>
    [Fact]
    public void AddHeading_AlsoDefinesTheNormalStyleItIsBasedOn()
    {
        using var stream = WriteAndValidate(body => body.AddHeading("Chapter", 1));
        var styles = ReadStylesXml(stream);

        Assert.Contains("w:styleId=\"Normal\"", styles, StringComparison.Ordinal);
        Assert.Contains("<w:basedOn w:val=\"Normal\" />", styles, StringComparison.Ordinal);
    }

    [Fact]
    public void AddHeading_UsedRepeatedly_DefinesEachStyleOnce()
    {
        using var stream = WriteAndValidate(body =>
        {
            body.AddHeading("First", 1);
            body.AddHeading("Second", 1);
            body.AddHeading("Third", 2);
        });

        var styles = ReadStylesXml(stream);

        Assert.Equal(1, CountOccurrences(styles, "w:styleId=\"Heading1\""));
        Assert.Equal(1, CountOccurrences(styles, "w:styleId=\"Heading2\""));
        Assert.Equal(1, CountOccurrences(styles, "w:styleId=\"Normal\""));
    }

    [Fact]
    public void AddParagraph_WithoutAStyle_DoesNotCreateAStylesPart()
    {
        using var stream = WriteAndValidate(body => body.AddParagraph("plain"));

        Assert.Equal(string.Empty, ReadStylesXml(stream));
    }

    /// <summary>
    /// An unknown style identifier belongs to the document's own template, so it is written through
    /// rather than replaced by an invented definition.
    /// </summary>
    [Fact]
    public void ApplyFormat_WithUnknownStyleId_WritesTheReferenceWithoutDefiningIt()
    {
        using var stream = WriteAndValidate(body => body.AddParagraph("x", new ParagraphFormat { StyleId = "CompanyQuote" }));

        Assert.Contains("w:val=\"CompanyQuote\"", ReadMainDocumentXml(stream), StringComparison.Ordinal);
        Assert.DoesNotContain("w:styleId=\"CompanyQuote\"", ReadStylesXml(stream), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(7)]
    public void AddHeading_WithLevelOutOfRange_Throws(int level)
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var body = document.GetBody();

        Assert.Throws<ArgumentOutOfRangeException>(() => body.AddHeading("x", level));
    }

    [Theory]
    [InlineData(0d)]
    [InlineData(-1d)]
    public void LineSpacing_NotPositive_ThrowsOnAssignment(double lineSpacing)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new ParagraphFormat { LineSpacing = lineSpacing });
    }

    [Fact]
    public void SpacingBefore_Negative_ThrowsOnAssignment()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new ParagraphFormat { SpacingBefore = -1 });
    }

    [Fact]
    public void IndentLeft_Negative_IsAllowed()
    {
        Assert.Equal(-12, new ParagraphFormat { IndentLeft = -12 }.IndentLeft);
    }

    [Fact]
    public void PageBreakBefore_RoundTrips()
    {
        using var stream = WriteAndValidate(body => body.AddParagraph("New page", new ParagraphFormat { PageBreakBefore = true }));

        using var document = OpenDocument(stream, isEditable: false);

        Assert.True(document.GetBody().Paragraphs[0].Format.PageBreakBefore);
    }

    [Fact]
    public void KeepWithNext_And_KeepLines_RoundTrip()
    {
        using var stream = WriteAndValidate(body =>
            body.AddParagraph("Stays together", new ParagraphFormat { KeepWithNext = true, KeepLines = true }));

        using var document = OpenDocument(stream, isEditable: false);
        var format = document.GetBody().Paragraphs[0].Format;

        Assert.True(format.KeepWithNext);
        Assert.True(format.KeepLines);
    }

    /// <summary>
    /// <c>CT_PPr</c> orders its children strictly: the keep and page-break switches come first, then
    /// numbering, then spacing and indentation, then justification. A format that sets all of them is
    /// where a hand-written order would break.
    /// </summary>
    [Fact]
    public void EveryProperty_TogetherIsSchemaValidAndRoundTrips()
    {
        var format = new ParagraphFormat
        {
            StyleId = WordStyleIds.Normal,
            Alignment = ParagraphAlignment.Justify,
            SpacingBefore = 6,
            SpacingAfter = 8,
            LineSpacing = 1.15,
            IndentLeft = 24,
            IndentRight = 12,
            IndentFirstLine = -12,
            PageBreakBefore = true,
            KeepWithNext = true,
            KeepLines = false,
            ListStyle = ListStyle.Number,
            ListLevel = 2,
        };

        using var stream = WriteAndValidate(body => body.AddParagraph("everything", format));

        using var document = OpenDocument(stream, isEditable: false);
        var actual = document.GetBody().Paragraphs[0].Format;

        Assert.Equal(format, actual);
    }

    [Fact]
    public void Merge_LetsTheArgumentWinAndKeepsTheRest()
    {
        var baseFormat = new ParagraphFormat { Alignment = ParagraphAlignment.Left, SpacingAfter = 6 };

        var merged = baseFormat.Merge(new ParagraphFormat { Alignment = ParagraphAlignment.Center });

        Assert.Equal(ParagraphAlignment.Center, merged.Alignment);
        Assert.Equal(6, merged.SpacingAfter);
        Assert.Equal(ParagraphAlignment.Left, baseFormat.Alignment);
    }

    private static int CountOccurrences(string haystack, string needle)
    {
        var count = 0;
        var index = haystack.IndexOf(needle, StringComparison.Ordinal);

        while (index >= 0)
        {
            count++;
            index = haystack.IndexOf(needle, index + needle.Length, StringComparison.Ordinal);
        }

        return count;
    }
}
