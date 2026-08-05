using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.Tests;

/// <summary>
/// Covers run-level formatting: what is written, what reads back, and what is rejected.
/// </summary>
public class TextFormattingTest : WordTestBase
{
    [Fact]
    public void AddText_WithFormat_RoundTripsEveryModelledProperty()
    {
        var format = new TextFormat
        {
            Bold = true,
            Italic = true,
            Underline = UnderlineType.Double,
            Strikethrough = true,
            FontName = "Georgia",
            FontSize = 13.5,
            Color = "#c00000",
        };

        using var stream = WriteAndValidate(body => body.AddParagraph().AddText("formatted", format));

        using var document = OpenDocument(stream, isEditable: false);
        var actual = document.GetBody().Paragraphs[0].Runs[0].Format;

        Assert.True(actual.Bold);
        Assert.True(actual.Italic);
        Assert.Equal(UnderlineType.Double, actual.Underline);
        Assert.True(actual.Strikethrough);
        Assert.Equal("Georgia", actual.FontName);
        Assert.Equal(13.5, actual.FontSize);
        Assert.Equal("C00000", actual.Color);
    }

    /// <summary>
    /// A half-point font size is the finest step WordprocessingML can store, so it has to survive the
    /// conversion rather than being rounded to a whole point.
    /// </summary>
    [Fact]
    public void FontSize_HalfPoint_IsStoredExactly()
    {
        using var stream = WriteAndValidate(body => body.AddParagraph().AddText("x", new TextFormat { FontSize = 10.5 }));

        Assert.Contains("w:val=\"21\"", ReadMainDocumentXml(stream), StringComparison.Ordinal);
    }

    /// <summary>
    /// Bold off is not the same as bold unspecified: the element has to be present with an explicit
    /// "off" value so that it overrides a bold paragraph style.
    /// </summary>
    [Fact]
    public void Bold_SetToFalse_IsWrittenExplicitly()
    {
        using var stream = WriteAndValidate(body => body.AddParagraph().AddText("x", new TextFormat { Bold = false }));

        var bold = ReadDocumentElement(stream, document => document.Descendants<WordLib.Bold>().Single());
        Assert.NotNull(bold.Val);
        Assert.False(bold.Val.Value);

        using var document = OpenDocument(stream, isEditable: false);
        Assert.False(document.GetBody().Paragraphs[0].Runs[0].Format.Bold);
    }

    [Fact]
    public void Bold_SetToTrue_UsesTheShorthandElement()
    {
        using var stream = WriteAndValidate(body => body.AddParagraph().AddText("x", new TextFormat { Bold = true }));

        Assert.Contains("<w:b />", ReadMainDocumentXml(stream), StringComparison.Ordinal);
    }

    /// <summary>
    /// Unspecified has to stay unspecified all the way through, otherwise a run would silently
    /// override the style it is supposed to inherit from.
    /// </summary>
    [Fact]
    public void AddText_WithoutFormat_LeavesEveryPropertyUnset()
    {
        using var stream = WriteAndValidate(body => body.AddParagraph("plain"));

        using var document = OpenDocument(stream, isEditable: false);
        var format = document.GetBody().Paragraphs[0].Runs[0].Format;

        Assert.True(format.IsEmpty);
        Assert.DoesNotContain("<w:rPr>", ReadMainDocumentXml(stream), StringComparison.Ordinal);
    }

    [Fact]
    public void ApplyFormat_OnExistingRun_KeepsPropertiesItDoesNotSet()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var run = document.GetBody().AddParagraph().AddRun("x", new TextFormat { Bold = true, FontSize = 12 });

        run.ApplyFormat(new TextFormat { Italic = true });

        Assert.True(run.Format.Bold);
        Assert.True(run.Format.Italic);
        Assert.Equal(12, run.Format.FontSize);
    }

    [Fact]
    public void RunText_Assigned_ReplacesContentAndKeepsFormatting()
    {
        using var stream = new MemoryStream();
        using var document = CreateDocument(stream);
        var run = document.GetBody().AddParagraph().AddRun("before", new TextFormat { Bold = true });

        run.Text = "after";

        Assert.Equal("after", run.Text);
        Assert.True(run.Format.Bold);
    }

    [Fact]
    public void Merge_LetsTheArgumentWinAndKeepsTheRest()
    {
        var baseFormat = new TextFormat { FontName = "Calibri", FontSize = 11, Bold = false };

        var merged = baseFormat.Merge(new TextFormat { Bold = true, Color = "FF0000" });

        Assert.True(merged.Bold);
        Assert.Equal("FF0000", merged.Color);
        Assert.Equal("Calibri", merged.FontName);
        Assert.Equal(11, merged.FontSize);
        Assert.False(baseFormat.Bold);
    }

    [Fact]
    public void Merge_WithNull_ReturnsTheSameFormat()
    {
        var format = new TextFormat { Bold = true };

        Assert.Same(format, format.Merge(null));
    }

    [Theory]
    [InlineData("FF0000", "FF0000")]
    [InlineData("#ff0000", "FF0000")]
    [InlineData(" 00ff00 ", "00FF00")]
    [InlineData("auto", "auto")]
    [InlineData("AUTO", "auto")]
    public void Color_IsNormalized(string input, string expected)
    {
        Assert.Equal(expected, new TextFormat { Color = input }.Color);
    }

    [Theory]
    [InlineData("red")]
    [InlineData("FF00")]
    [InlineData("GGGGGG")]
    [InlineData("")]
    public void Color_Invalid_ThrowsOnAssignment(string input)
    {
        Assert.Throws<ArgumentException>(() => new TextFormat { Color = input });
    }

    /// <summary>
    /// An ARGB value is a likely copy/paste from the Excel module, where colours do carry alpha.
    /// Refusing it beats discarding the alpha channel without telling anyone.
    /// </summary>
    [Fact]
    public void Color_WithAlphaChannel_IsRejectedWithAnExplanation()
    {
        var exception = Assert.Throws<ArgumentException>(() => new TextFormat { Color = "FFFF0000" });

        Assert.Contains("alpha", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(0d)]
    [InlineData(-1d)]
    [InlineData(2000d)]
    [InlineData(double.NaN)]
    public void FontSize_OutOfRange_ThrowsOnAssignment(double size)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new TextFormat { FontSize = size });
    }

    [Fact]
    public void FontName_Blank_ThrowsOnAssignment()
    {
        Assert.Throws<ArgumentException>(() => new TextFormat { FontName = "  " });
    }

    [Fact]
    public void IsEmpty_IsTrueOnlyWhenNothingIsSet()
    {
        Assert.True(new TextFormat().IsEmpty);
        Assert.False(new TextFormat { Bold = false }.IsEmpty);
    }

    [Fact]
    public void Highlight_RoundTrips()
    {
        using var stream = WriteAndValidate(body =>
            body.AddParagraph().AddText("marked", new TextFormat { Highlight = HighlightColor.Yellow }));

        using var document = OpenDocument(stream, isEditable: false);

        Assert.Equal(HighlightColor.Yellow, document.GetBody().Paragraphs[0].Runs[0].Format.Highlight);
    }

    /// <summary>
    /// The highlight palette is fixed by the format, so every value has to map to a real one rather
    /// than being silently dropped.
    /// </summary>
    [Theory]
    [InlineData(HighlightColor.None)]
    [InlineData(HighlightColor.Yellow)]
    [InlineData(HighlightColor.Cyan)]
    [InlineData(HighlightColor.DarkMagenta)]
    [InlineData(HighlightColor.LightGray)]
    [InlineData(HighlightColor.White)]
    public void Highlight_EveryValueRoundTrips(HighlightColor highlight)
    {
        using var stream = WriteAndValidate(body =>
            body.AddParagraph().AddText("x", new TextFormat { Highlight = highlight }));

        using var document = OpenDocument(stream, isEditable: false);

        Assert.Equal(highlight, document.GetBody().Paragraphs[0].Runs[0].Format.Highlight);
    }

    [Theory]
    [InlineData(TextVerticalPosition.Superscript)]
    [InlineData(TextVerticalPosition.Subscript)]
    [InlineData(TextVerticalPosition.Baseline)]
    public void VerticalPosition_RoundTrips(TextVerticalPosition position)
    {
        using var stream = WriteAndValidate(body =>
            body.AddParagraph().AddText("2", new TextFormat { VerticalPosition = position }));

        using var document = OpenDocument(stream, isEditable: false);

        Assert.Equal(position, document.GetBody().Paragraphs[0].Runs[0].Format.VerticalPosition);
    }

    /// <summary>
    /// All-caps is a rendering instruction, so the stored characters keep their original case and the
    /// text still reads back as it was written.
    /// </summary>
    [Fact]
    public void AllCaps_ChangesRenderingWithoutChangingTheText()
    {
        using var stream = WriteAndValidate(body =>
            body.AddParagraph().AddText("Heading text", new TextFormat { AllCaps = true }));

        using var document = OpenDocument(stream, isEditable: false);
        var run = document.GetBody().Paragraphs[0].Runs[0];

        Assert.True(run.Format.AllCaps);
        Assert.Equal("Heading text", run.Text);
    }

    [Fact]
    public void SmallCaps_RoundTrips()
    {
        using var stream = WriteAndValidate(body =>
            body.AddParagraph().AddText("x", new TextFormat { SmallCaps = true }));

        using var document = OpenDocument(stream, isEditable: false);

        Assert.True(document.GetBody().Paragraphs[0].Runs[0].Format.SmallCaps);
    }

    /// <summary>
    /// Superscript and subscript sit in a strict position within <c>CT_RPr</c>, well after the toggles.
    /// A combined format is where a hand-written child order would break.
    /// </summary>
    [Fact]
    public void EveryProperty_TogetherIsSchemaValidAndRoundTrips()
    {
        var format = new TextFormat
        {
            Bold = true,
            Italic = true,
            Underline = UnderlineType.Wave,
            Strikethrough = true,
            AllCaps = true,
            SmallCaps = false,
            Highlight = HighlightColor.Cyan,
            VerticalPosition = TextVerticalPosition.Superscript,
            FontName = "Verdana",
            FontSize = 9.5,
            Color = "112233",
        };

        using var stream = WriteAndValidate(body => body.AddParagraph().AddText("everything", format));

        using var document = OpenDocument(stream, isEditable: false);
        var actual = document.GetBody().Paragraphs[0].Runs[0].Format;

        Assert.Equal(format, actual);
    }
}
