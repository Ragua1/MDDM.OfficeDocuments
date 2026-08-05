using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Styles;
using OfficeDocuments.Excel.TestKit;
using Color = System.Drawing.Color;

namespace OfficeDocuments.Excel.IntegrationTests.Styles;

/// <summary>
/// The identity contract of the stylesheet: equal input must reuse an entry, differing input must
/// allocate a new one, and N distinct styles must produce exactly N entries.
/// </summary>
/// <remarks>
/// This matters beyond tidiness — a stylesheet that grows one entry per call is the O(N²) cost the
/// verdict document flags, and duplicate entries are how a workbook quietly triples in size.
/// </remarks>
public class StyleDedupTests : SpreadsheetTestBase
{
    [Fact]
    public void SameFontTwice_ReusesTheSameEntry()
    {
        using var workbook = CreateInMemorySpreadsheet();

        var first = workbook.CreateStyle(new Font { Bold = true, FontSize = 14 });
        var countAfterFirst = StylesheetProbe.FontCount(first);
        var second = workbook.CreateStyle(new Font { Bold = true, FontSize = 14 });

        Assert.Equal(first.FontId, second.FontId);
        Assert.Equal(first.StyleIndex, second.StyleIndex);
        Assert.Equal(countAfterFirst, StylesheetProbe.FontCount(second));
    }

    [Fact]
    public void FontDifferingInOneAttribute_AllocatesANewEntry()
    {
        using var workbook = CreateInMemorySpreadsheet();

        var bold = workbook.CreateStyle(new Font { Bold = true, FontSize = 14 });
        var italic = workbook.CreateStyle(new Font { Italic = true, FontSize = 14 });

        Assert.NotEqual(bold.FontId, italic.FontId);
        Assert.NotEqual(bold.StyleIndex, italic.StyleIndex);
    }

    [Fact]
    public void DistinctFonts_ProduceExactlyOneEntryEach()
    {
        using var workbook = CreateInMemorySpreadsheet();

        var baseline = workbook.CreateStyle();
        var fontsBefore = StylesheetProbe.FontCount(baseline);

        var sizes = new[] { 8d, 9d, 10d, 12d, 16d, 24d };
        var styles = sizes.Select(size => workbook.CreateStyle(new Font { FontSize = size })).ToList();

        Assert.Equal(sizes.Length, styles.Select(style => style.FontId).Distinct().Count());
        Assert.Equal(fontsBefore + sizes.Length, StylesheetProbe.FontCount(styles[0]));
    }

    [Fact]
    public void RepeatingTheSameStyleManyTimes_DoesNotGrowTheStylesheet()
    {
        using var workbook = CreateInMemorySpreadsheet();

        var first = workbook.CreateStyle(
            font: new Font { Bold = true },
            fill: new Fill(Color.Yellow),
            border: new Border(BorderStyleValues.Thin));

        var fonts = StylesheetProbe.FontCount(first);
        var fills = StylesheetProbe.FillCount(first);
        var borders = StylesheetProbe.BorderCount(first);
        var cellFormats = StylesheetProbe.CellFormatCount(first);

        for (var i = 0; i < 50; i++)
        {
            var repeat = workbook.CreateStyle(
                font: new Font { Bold = true },
                fill: new Fill(Color.Yellow),
                border: new Border(BorderStyleValues.Thin));

            Assert.Equal(first.StyleIndex, repeat.StyleIndex);
        }

        Assert.Equal(fonts, StylesheetProbe.FontCount(first));
        Assert.Equal(fills, StylesheetProbe.FillCount(first));
        Assert.Equal(borders, StylesheetProbe.BorderCount(first));
        Assert.Equal(cellFormats, StylesheetProbe.CellFormatCount(first));
    }

    [Fact]
    public void BorderMatchingTheDefaultEntry_ReusesIndexZeroInsteadOfDuplicatingIt()
    {
        // The default stylesheet starts with an empty <border/> at index 0. Asking for an empty
        // border must resolve to that entry, not append a second identical one.
        using var workbook = CreateInMemorySpreadsheet();

        var baseline = workbook.CreateStyle();
        var bordersBefore = StylesheetProbe.BorderCount(baseline);

        var style = workbook.CreateStyle(border: new Border());

        Assert.Equal(0, style.BorderId);
        Assert.Equal(bordersBefore, StylesheetProbe.BorderCount(style));
    }

    [Fact]
    public void FontMatchingTheDefaultEntry_ReusesIndexZeroInsteadOfDuplicatingIt()
    {
        // The default font is 11pt black Calibri; requesting exactly that must not append a copy.
        using var workbook = CreateInMemorySpreadsheet();

        var baseline = workbook.CreateStyle();
        var fontsBefore = StylesheetProbe.FontCount(baseline);

        var style = workbook.CreateStyle(new Font
        {
            FontSize = 11,
            Color = Color.Black,
            FontName = FontNameValues.Calibri
        });

        Assert.Equal(0, style.FontId);
        Assert.Equal(fontsBefore, StylesheetProbe.FontCount(style));
    }

    [Fact]
    public void SameCustomNumberFormatTwice_ReusesTheAllocatedId()
    {
        using var workbook = CreateInMemorySpreadsheet();

        var first = workbook.CreateStyle(numberFormat: new NumberingFormat("0.0000"));
        var second = workbook.CreateStyle(numberFormat: new NumberingFormat("0.0000"));

        Assert.Equal(first.NumberFormatId, second.NumberFormatId);
        Assert.Equal(1, StylesheetProbe.NumberingFormatCount(second));
    }

    [Fact]
    public void DistinctCustomNumberFormats_GetConsecutiveIds()
    {
        using var workbook = CreateInMemorySpreadsheet();

        var first = workbook.CreateStyle(numberFormat: new NumberingFormat("0.0"));
        var second = workbook.CreateStyle(numberFormat: new NumberingFormat("0.00000"));

        Assert.Equal(170, first.NumberFormatId);
        Assert.Equal(171, second.NumberFormatId);
        Assert.Equal(2, StylesheetProbe.NumberingFormatCount(second));
    }

    [Fact]
    public void SameFacetsDifferentAlignment_AllocateDifferentCellFormats()
    {
        using var workbook = CreateInMemorySpreadsheet();

        var left = workbook.CreateStyle(
            font: new Font { Bold = true },
            alignment: new Alignment { Horizontal = HorizontalAlignmentValues.Left });
        var right = workbook.CreateStyle(
            font: new Font { Bold = true },
            alignment: new Alignment { Horizontal = HorizontalAlignmentValues.Right });

        Assert.Equal(left.FontId, right.FontId);
        Assert.NotEqual(left.StyleIndex, right.StyleIndex);
    }

    [Fact]
    public void StyleWithNoFacets_ResolvesToTheDefaultCellFormat()
    {
        using var workbook = CreateInMemorySpreadsheet();

        var style = workbook.CreateStyle();

        Assert.Equal(0U, style.StyleIndex);
        Assert.Equal(0, style.FontId);
        Assert.Equal(0, style.FillId);
        Assert.Equal(0, style.BorderId);
        Assert.Equal(0, style.NumberFormatId);
    }
}
