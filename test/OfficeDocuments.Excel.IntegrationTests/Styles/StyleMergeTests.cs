using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Interfaces;
using OfficeDocuments.Excel.Styles;
using OfficeDocuments.Excel.TestKit;
using Color = System.Drawing.Color;

namespace OfficeDocuments.Excel.IntegrationTests.Styles;

/// <summary>
/// The merge truth table for <see cref="IStyle.CreateMergedStyle"/>.
/// </summary>
/// <remarks>
/// Style composition is the library's headline value-add over the raw SDK, but its semantics were
/// previously pinned by four examples rather than by rule. Each facet is exercised across the four
/// combinations of "base carries it" × "overlay carries it".
/// </remarks>
public class StyleMergeTests : SpreadsheetTestBase
{
    // ---- font ----------------------------------------------------------------------------

    [Fact]
    public void Font_OverlayOnly_TakesTheOverlayOnTopOfTheWorkbookDefault()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var baseStyle = workbook.CreateStyle(border: new Border(BorderStyleValues.Thin));
        var overlay = workbook.CreateStyle(new Font { Bold = true });

        var merged = baseStyle.CreateMergedStyle(overlay);

        // The base has no font of its own, so the merge runs against the default font entry and
        // the result carries the workbook defaults as well as the overlay's bold. That means it
        // is a new entry rather than a reuse of the overlay's.
        Assert.NotEqual(overlay.FontId, merged.FontId);
        OoxmlAssert.RendersAs(
            """
            <font>
              <b val="1" /><sz val="11" /><color rgb="FF000000" /><name val="Calibri" />
            </font>
            """,
            StylesheetProbe.Font(merged));
    }

    [Fact]
    public void Font_BaseOnly_KeepsTheBase()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var baseStyle = workbook.CreateStyle(new Font { Bold = true });
        var overlay = workbook.CreateStyle(border: new Border(BorderStyleValues.Thin));

        var merged = baseStyle.CreateMergedStyle(overlay);

        Assert.Equal(baseStyle.FontId, merged.FontId);
    }

    [Fact]
    public void Font_BothSides_MergePerAttributeWithOverlayWinning()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var baseStyle = workbook.CreateStyle(new Font { FontSize = 20, Italic = true });
        var overlay = workbook.CreateStyle(new Font { FontSize = 8, Bold = true });

        var merged = baseStyle.CreateMergedStyle(overlay);

        // Size is overridden by the overlay; italic survives from the base; bold arrives from the
        // overlay — and the result still honours the CT_Font sequence.
        OoxmlAssert.RendersAs(
            """<font><b val="1" /><i val="1" /><sz val="8" /></font>""",
            StylesheetProbe.Font(merged));
    }

    [Fact]
    public void Font_NeitherSide_StaysDefault()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var baseStyle = workbook.CreateStyle(fill: new Fill(Color.Yellow));
        var overlay = workbook.CreateStyle(border: new Border(BorderStyleValues.Thin));

        var merged = baseStyle.CreateMergedStyle(overlay);

        Assert.Equal(0, merged.FontId);
    }

    // ---- fill ----------------------------------------------------------------------------

    [Fact]
    public void Fill_OverlayOnly_TakesTheOverlay()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var baseStyle = workbook.CreateStyle(new Font { Bold = true });
        var overlay = workbook.CreateStyle(fill: new Fill(Color.Yellow));

        var merged = baseStyle.CreateMergedStyle(overlay);

        Assert.Equal(overlay.FillId, merged.FillId);
    }

    [Fact]
    public void Fill_BothSides_OverlayColorWins()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var baseStyle = workbook.CreateStyle(fill: new Fill(Color.Yellow));
        var overlay = workbook.CreateStyle(fill: new Fill(Color.Red));

        var merged = baseStyle.CreateMergedStyle(overlay);

        OoxmlAssert.RendersAs(
            """<fill><patternFill patternType="solid"><fgColor rgb="FFFF0000" /></patternFill></fill>""",
            StylesheetProbe.Fill(merged));
    }

    // ---- border --------------------------------------------------------------------------

    [Fact]
    public void Border_BothSides_MergePerEdge()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var baseStyle = workbook.CreateStyle(border: new Border { Left = BorderStyleValues.Thin });
        var overlay = workbook.CreateStyle(border: new Border { Top = BorderStyleValues.Medium });

        var merged = baseStyle.CreateMergedStyle(overlay);

        OoxmlAssert.RendersAs(
            """<border><left style="thin" /><top style="medium" /></border>""",
            StylesheetProbe.Border(merged));
        OoxmlAssert.ChildOrder(StylesheetProbe.Border(merged), "left", "top");
    }

    [Fact]
    public void Border_SameEdgeOnBothSides_OverlayWins()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var baseStyle = workbook.CreateStyle(border: new Border { Left = BorderStyleValues.Thin });
        var overlay = workbook.CreateStyle(border: new Border { Left = BorderStyleValues.Thick });

        var merged = baseStyle.CreateMergedStyle(overlay);

        OoxmlAssert.RendersAs("""<border><left style="thick" /></border>""", StylesheetProbe.Border(merged));
    }

    // ---- number format -------------------------------------------------------------------

    [Fact]
    public void NumberFormat_OverlayOnly_TakesTheOverlay()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var baseStyle = workbook.CreateStyle(new Font { Bold = true });
        var overlay = workbook.CreateStyle(numberFormat: new NumberingFormat("0.00#"));

        var merged = baseStyle.CreateMergedStyle(overlay);

        Assert.Equal(overlay.NumberFormatId, merged.NumberFormatId);
    }

    [Fact]
    public void NumberFormat_BothSides_OverlayReplacesRatherThanMerges()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var baseStyle = workbook.CreateStyle(numberFormat: new NumberingFormat("0.0#"));
        var overlay = workbook.CreateStyle(numberFormat: new NumberingFormat("0.000#"));

        var merged = baseStyle.CreateMergedStyle(overlay);

        Assert.Equal(overlay.NumberFormatId, merged.NumberFormatId);
        Assert.Equal("0.000#", StylesheetProbe.NumberingFormat(merged)?.FormatCode?.Value);
    }

    [Fact]
    public void NumberFormat_BuiltInOverlay_IsReferencedByItsBuiltInId()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var baseStyle = workbook.CreateStyle(new Font { Bold = true });
        var overlay = workbook.CreateStyle(numberFormat: new NumberingFormat("@"));

        var merged = baseStyle.CreateMergedStyle(overlay);

        Assert.Equal(49, merged.NumberFormatId);
    }

    // ---- alignment -----------------------------------------------------------------------

    [Fact]
    public void Alignment_OverlayOnly_TakesTheOverlay()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var baseStyle = workbook.CreateStyle(new Font { Bold = true });
        var overlay = workbook.CreateStyle(alignment: new Alignment { Horizontal = HorizontalAlignmentValues.Center });

        var merged = baseStyle.CreateMergedStyle(overlay);

        OoxmlAssert.RendersAs("""<alignment horizontal="center" />""", StylesheetProbe.CellFormat(merged).Alignment!);
    }

    [Fact]
    public void Alignment_BothSides_BaseWinsBecauseAlignmentIsNotMerged()
    {
        // Documented behaviour, not an accident: Style.CreateMergedStyle only adopts the overlay's
        // alignment when the base has none.
        using var workbook = CreateInMemorySpreadsheet();
        var baseStyle = workbook.CreateStyle(alignment: new Alignment { Horizontal = HorizontalAlignmentValues.Left });
        var overlay = workbook.CreateStyle(alignment: new Alignment { Vertical = VerticalAlignmentValues.Bottom });

        var merged = baseStyle.CreateMergedStyle(overlay);

        OoxmlAssert.RendersAs("""<alignment horizontal="left" />""", StylesheetProbe.CellFormat(merged).Alignment!);
    }

    [Fact]
    public void Alignment_NeitherSide_StaysAbsent()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var baseStyle = workbook.CreateStyle(new Font { Bold = true });
        var overlay = workbook.CreateStyle(fill: new Fill(Color.Yellow));

        var merged = baseStyle.CreateMergedStyle(overlay);

        Assert.Null(StylesheetProbe.CellFormat(merged).Alignment);
    }

    // ---- whole-style behaviour -------------------------------------------------------------

    [Fact]
    public void MergeWithNull_ReturnsTheBaseUnchanged()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var baseStyle = workbook.CreateStyle(new Font { Bold = true });

        var merged = baseStyle.CreateMergedStyle(null);

        Assert.Same(baseStyle, merged);
    }

    [Fact]
    public void MergeIsNotCommutative_AndBothDirectionsAreStable()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var small = workbook.CreateStyle(new Font { FontSize = 8 });
        var large = workbook.CreateStyle(new Font { FontSize = 20 });

        var smallOverLarge = large.CreateMergedStyle(small);
        var largeOverSmall = small.CreateMergedStyle(large);

        Assert.NotEqual(smallOverLarge.FontId, largeOverSmall.FontId);
        OoxmlAssert.RendersAs("""<font><sz val="8" /></font>""", StylesheetProbe.Font(smallOverLarge));
        OoxmlAssert.RendersAs("""<font><sz val="20" /></font>""", StylesheetProbe.Font(largeOverSmall));
    }

    [Fact]
    public void MergingTheSamePairTwice_ReusesTheSameCellFormat()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var baseStyle = workbook.CreateStyle(new Font { FontSize = 20 });
        var overlay = workbook.CreateStyle(fill: new Fill(Color.Yellow));

        var first = baseStyle.CreateMergedStyle(overlay);
        var cellFormatsAfterFirst = StylesheetProbe.CellFormatCount(first);
        var second = baseStyle.CreateMergedStyle(overlay);

        Assert.Equal(first.StyleIndex, second.StyleIndex);
        Assert.Equal(cellFormatsAfterFirst, StylesheetProbe.CellFormatCount(second));
    }

    // ---- across workbooks ------------------------------------------------------------------

    [Fact]
    public void MergeAcrossWorkbooks_CopiesEveryFacetIntoTheTargetStylesheet()
    {
        // The existing coverage checked font and fill only. A facet that is not copied would leave
        // the merged style pointing at an index that means something else in the target workbook.
        using var target = CreateInMemorySpreadsheet();
        using var source = CreateInMemorySpreadsheet();

        var targetStyle = target.CreateStyle(new Font { FontSize = 9 });
        var sourceStyle = source.CreateStyle(
            font: new Font { Bold = true },
            fill: new Fill(Color.Aqua),
            border: new Border { Left = BorderStyleValues.Thick },
            numberFormat: new NumberingFormat("0.0000#"),
            alignment: new Alignment { Horizontal = HorizontalAlignmentValues.Center });

        var merged = targetStyle.CreateMergedStyle(sourceStyle);

        // Every id must resolve inside the target stylesheet and carry the source's content.
        Assert.True(StylesheetProbe.Font(merged).Bold?.Val?.Value);
        Assert.Equal("FF00FFFF", StylesheetProbe.Fill(merged).PatternFill?.ForegroundColor?.Rgb?.Value);
        OoxmlAssert.RendersAs("""<border><left style="thick" /></border>""", StylesheetProbe.Border(merged));
        Assert.Equal("0.0000#", StylesheetProbe.NumberingFormat(merged)?.FormatCode?.Value);
        OoxmlAssert.RendersAs("""<alignment horizontal="center" />""", StylesheetProbe.CellFormat(merged).Alignment!);
    }

    [Fact]
    public void MergeAcrossWorkbooks_DoesNotTouchTheSourceStylesheet()
    {
        using var target = CreateInMemorySpreadsheet();
        using var source = CreateInMemorySpreadsheet();

        var targetStyle = target.CreateStyle(new Font { FontSize = 9 });
        var sourceStyle = source.CreateStyle(fill: new Fill(Color.Aqua));

        var fillsInSourceBefore = StylesheetProbe.FillCount(sourceStyle);
        targetStyle.CreateMergedStyle(sourceStyle);

        Assert.Equal(fillsInSourceBefore, StylesheetProbe.FillCount(sourceStyle));
    }

    [Fact]
    public void ChainedMerges_AccumulateEveryFacet()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var font = workbook.CreateStyle(new Font { Bold = true });
        var fill = workbook.CreateStyle(fill: new Fill(Color.Yellow));
        var border = workbook.CreateStyle(border: new Border(BorderStyleValues.Thin));

        var merged = font.CreateMergedStyle(fill).CreateMergedStyle(border);

        Assert.Equal(font.FontId, merged.FontId);
        Assert.Equal(fill.FillId, merged.FillId);
        Assert.Equal(border.BorderId, merged.BorderId);
    }
}
