using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Styles;
using OfficeDocuments.Excel.TestKit;
using Color = System.Drawing.Color;

namespace OfficeDocuments.Excel.IntegrationTests.Styles;

/// <summary>
/// Asserts what each style facet actually renders into <c>styles.xml</c>.
/// </summary>
/// <remarks>
/// The pre-existing style tests only checked that an id greater than zero was handed out, which
/// cannot distinguish blue from red, bold from italic, or a correctly ordered font from one Excel
/// will reject. These pin the produced XML instead.
/// </remarks>
public class StyleRenderingTests : SpreadsheetTestBase
{
    [Fact]
    public void Font_Bold_RendersBoldElement()
    {
        using var workbook = CreateInMemorySpreadsheet();

        var style = workbook.CreateStyle(new Font { Bold = true });

        OoxmlAssert.RendersAs("""<font><b val="1" /></font>""", StylesheetProbe.Font(style));
    }

    [Fact]
    public void Font_AllFacets_RenderInSchemaOrder()
    {
        using var workbook = CreateInMemorySpreadsheet();

        var style = workbook.CreateStyle(new Font
        {
            FontName = FontNameValues.Arial,
            FontSize = 20,
            Bold = true,
            Italic = true,
            Color = Color.DarkBlue,
            Underline = UnderlineValues.Double
        });

        var font = StylesheetProbe.Font(style);

        // Written deliberately in a different order from CT_Font; the library must still emit
        // b, i, u, sz, color, name.
        OoxmlAssert.ChildOrder(font, "b", "i", "u", "sz", "color", "name");
        OoxmlAssert.RendersAs(
            """
            <font>
              <b val="1" /><i val="1" /><u val="double" />
              <sz val="20" /><color rgb="FF00008B" /><name val="Arial" />
            </font>
            """,
            font);
    }

    [Theory]
    [InlineData(UnderlineValues.None, "none")]
    [InlineData(UnderlineValues.Single, "single")]
    [InlineData(UnderlineValues.Double, "double")]
    [InlineData(UnderlineValues.SingleAccounting, "singleAccounting")]
    [InlineData(UnderlineValues.DoubleAccounting, "doubleAccounting")]
    public void Font_Underline_RendersEveryVariant(UnderlineValues underline, string expected)
    {
        using var workbook = CreateInMemorySpreadsheet();

        var style = workbook.CreateStyle(new Font { Underline = underline });

        OoxmlAssert.RendersAs($"""<font><u val="{expected}" /></font>""", StylesheetProbe.Font(style));
    }

    [Fact]
    public void Font_Color_RendersArgbNotRgb()
    {
        using var workbook = CreateInMemorySpreadsheet();

        var style = workbook.CreateStyle(new Font { Color = Color.Red });

        OoxmlAssert.RendersAs("""<font><color rgb="FFFF0000" /></font>""", StylesheetProbe.Font(style));
    }

    [Fact]
    public void Fill_ForegroundOnly_RendersSolidPatternFill()
    {
        using var workbook = CreateInMemorySpreadsheet();

        var style = workbook.CreateStyle(fill: new Fill(Color.Yellow));

        OoxmlAssert.RendersAs(
            """<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00" /></patternFill></fill>""",
            StylesheetProbe.Fill(style));
    }

    [Fact]
    public void Fill_BackgroundAndForeground_RenderBothColors()
    {
        using var workbook = CreateInMemorySpreadsheet();

        var style = workbook.CreateStyle(fill: new Fill(Color.Blue, Color.White));

        // CT_PatternFill declares fgColor before bgColor.
        var fill = StylesheetProbe.Fill(style);
        OoxmlAssert.ChildOrder(fill.PatternFill!, "fgColor", "bgColor");
        OoxmlAssert.RendersAs(
            """
            <fill>
              <patternFill patternType="solid">
                <fgColor rgb="FFFFFFFF" /><bgColor rgb="FF0000FF" />
              </patternFill>
            </fill>
            """,
            fill);
    }

    [Theory]
    [InlineData(BorderStyleValues.Thin, "thin")]
    [InlineData(BorderStyleValues.Medium, "medium")]
    [InlineData(BorderStyleValues.Thick, "thick")]
    [InlineData(BorderStyleValues.Double, "double")]
    [InlineData(BorderStyleValues.Dashed, "dashed")]
    [InlineData(BorderStyleValues.Dotted, "dotted")]
    [InlineData(BorderStyleValues.Hair, "hair")]
    [InlineData(BorderStyleValues.MediumDashed, "mediumDashed")]
    [InlineData(BorderStyleValues.DashDot, "dashDot")]
    public void Border_EveryStyleValue_RendersItsOoxmlName(BorderStyleValues borderStyle, string expected)
    {
        using var workbook = CreateInMemorySpreadsheet();

        var style = workbook.CreateStyle(border: new Border { Left = borderStyle });

        OoxmlAssert.RendersAs($"""<border><left style="{expected}" /></border>""", StylesheetProbe.Border(style));
    }

    [Fact]
    public void Border_EachEdgeIndependently_RendersOnlyThatEdge()
    {
        using var workbook = CreateInMemorySpreadsheet();

        var top = workbook.CreateStyle(border: new Border { Top = BorderStyleValues.Thin });
        var bottom = workbook.CreateStyle(border: new Border { Bottom = BorderStyleValues.Thin });

        OoxmlAssert.RendersAs("""<border><top style="thin" /></border>""", StylesheetProbe.Border(top));
        OoxmlAssert.RendersAs("""<border><bottom style="thin" /></border>""", StylesheetProbe.Border(bottom));
    }

    [Fact]
    public void Border_AllEdges_RenderInSchemaOrder()
    {
        using var workbook = CreateInMemorySpreadsheet();

        var style = workbook.CreateStyle(border: new Border(BorderStyleValues.Medium));

        // CT_Border declares left, right, top, bottom — the setter assigns them in a different order.
        OoxmlAssert.ChildOrder(StylesheetProbe.Border(style), "left", "right", "top", "bottom");
    }

    [Fact]
    public void Alignment_RendersOnTheCellFormatNotTheFont()
    {
        using var workbook = CreateInMemorySpreadsheet();

        var style = workbook.CreateStyle(alignment: new Alignment
        {
            Horizontal = HorizontalAlignmentValues.Center,
            Vertical = VerticalAlignmentValues.Center,
            WrapText = true
        });

        var alignment = StylesheetProbe.CellFormat(style).Alignment;

        Assert.NotNull(alignment);
        OoxmlAssert.RendersAs(
            """<alignment horizontal="center" vertical="center" wrapText="1" />""",
            alignment);
    }

    [Fact]
    public void NumberFormat_Custom_RendersFormatCodeAtTheAllocatedId()
    {
        using var workbook = CreateInMemorySpreadsheet();

        var style = workbook.CreateStyle(numberFormat: new NumberingFormat("#,##0.00#"));
        var numberFormat = StylesheetProbe.NumberingFormat(style);

        Assert.NotNull(numberFormat);
        Assert.Equal("#,##0.00#", numberFormat.FormatCode?.Value);
        Assert.Equal(170, style.NumberFormatId);
    }

    [Fact]
    public void NumberFormat_BuiltIn_IsReferencedByIdWithoutAddingAnEntry()
    {
        using var workbook = CreateInMemorySpreadsheet();

        var style = workbook.CreateStyle(numberFormat: new NumberingFormat("@"));

        Assert.Equal(49, style.NumberFormatId);
        Assert.Equal(0, StylesheetProbe.NumberingFormatCount(style));
    }

    [Fact]
    public void NumberFormat_CodeContainingQuotesAndBrackets_SurvivesVerbatim()
    {
        using var workbook = CreateInMemorySpreadsheet();

        const string formatCode = "#,##0.00 \"Kč\";[Red]-#,##0.00 \"Kč\"";
        var style = workbook.CreateStyle(numberFormat: new NumberingFormat(formatCode));

        Assert.Equal(formatCode, StylesheetProbe.NumberingFormat(style)?.FormatCode?.Value);
    }

    [Fact]
    public void CellFormat_CombiningEveryFacet_PointsAtEachAllocatedEntry()
    {
        using var workbook = CreateInMemorySpreadsheet(out var stream);

        var style = workbook.CreateStyle(
            font: new Font { Bold = true, Color = Color.White },
            fill: new Fill(Color.DarkSlateBlue),
            border: new Border(BorderStyleValues.Thin),
            numberFormat: new NumberingFormat("0.000"),
            alignment: new Alignment { Horizontal = HorizontalAlignmentValues.Right });

        var worksheet = workbook.AddWorksheet("Styled");
        worksheet.AddCell(1, 1, "value", style);

        OoxmlAssert.RendersAs("""<font><b val="1" /><color rgb="FFFFFFFF" /></font>""", StylesheetProbe.Font(style));
        OoxmlAssert.RendersAs(
            """<fill><patternFill patternType="solid"><fgColor rgb="FF483D8B" /></patternFill></fill>""",
            StylesheetProbe.Fill(style));
        OoxmlAssert.ChildOrder(StylesheetProbe.Border(style), "left", "right", "top", "bottom");
        Assert.Equal("0.000", StylesheetProbe.NumberingFormat(style)?.FormatCode?.Value);

        workbook.Close();
        SaveArtifact(stream, "every-style-facet.xlsx");
    }
}
