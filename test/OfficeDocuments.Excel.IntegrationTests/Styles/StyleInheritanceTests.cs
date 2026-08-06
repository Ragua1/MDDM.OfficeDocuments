using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Interfaces;
using OfficeDocuments.Excel.Styles;
using OfficeDocuments.Excel.TestKit;
using Color = System.Drawing.Color;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.IntegrationTests.Styles;

/// <summary>
/// Precedence along the sheet → row → cell chain.
/// </summary>
/// <remarks>
/// <para>
/// Only two tests covered this before and neither exercised a conflict, so what happens when two
/// levels set the same facet — or different ones — was unspecified in practice.
/// </para>
/// <para>
/// Two consequences of the composition model are worth stating, because both surprise on first
/// contact. First, every level composes by merging, so a style that reaches a cell through the
/// chain is a <em>new</em> stylesheet entry: comparing <c>StyleIndex</c> against the style handed
/// to <c>AddRow</c> does not work, and these tests assert on the resolved font instead. Second,
/// any merge that involves a level with no font of its own folds in the workbook default font
/// (11pt black Calibri), so the resolved font carries <c>sz</c>/<c>color</c>/<c>name</c> even when
/// the caller only asked for bold.
/// </para>
/// </remarks>
public class StyleInheritanceTests : SpreadsheetTestBase
{
    private static SpreadsheetLib.Font ResolvedFont(IStyle? style)
    {
        Assert.NotNull(style);

        return StylesheetProbe.Font(style);
    }

    private static SpreadsheetLib.Fill ResolvedFill(IStyle? style)
    {
        Assert.NotNull(style);

        return StylesheetProbe.Fill(style);
    }

    private static SpreadsheetLib.Border ResolvedBorder(IStyle? style)
    {
        Assert.NotNull(style);

        return StylesheetProbe.Border(style);
    }

    private static string? ForegroundColor(IStyle? style) =>
        ResolvedFill(style).PatternFill?.ForegroundColor?.Rgb?.Value;

    [Fact]
    public void CellWithoutOwnStyle_InheritsTheRowStyle()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var worksheet = workbook.AddWorksheet("Sheet1");
        var row = worksheet.AddRow(workbook.CreateStyle(new Font { Color = Color.DarkGoldenrod }));

        var cell = row.AddCell();

        Assert.NotNull(row.Style);
        Assert.NotNull(cell.Style);
        Assert.Equal(row.Style.StyleIndex, cell.Style.StyleIndex);
        Assert.Equal("FFB8860B", ResolvedFont(cell.Style).Color?.Rgb?.Value);
    }

    [Fact]
    public void CellWithoutOwnStyle_InheritsTheSheetStyle()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var worksheet = workbook.AddWorksheet("Sheet1", workbook.CreateStyle(new Font { Color = Color.DarkGoldenrod }));

        var cell = worksheet.AddCell();

        Assert.NotNull(worksheet.Style);
        Assert.NotNull(cell.Style);
        Assert.Equal(worksheet.Style.StyleIndex, cell.Style.StyleIndex);
        Assert.Equal("FFB8860B", ResolvedFont(cell.Style).Color?.Rgb?.Value);
    }

    [Fact]
    public void RowStyle_TakesPrecedenceOverSheetStyleForTheSameFacet()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var worksheet = workbook.AddWorksheet("Sheet1", workbook.CreateStyle(new Font { FontSize = 8 }));

        var cell = worksheet.AddRow(workbook.CreateStyle(new Font { FontSize = 20 })).AddCell();

        Assert.Equal(20d, ResolvedFont(cell.Style).FontSize?.Val?.Value);
    }

    [Fact]
    public void DifferentFacetsAtSheetAndRowLevel_BothSurviveOnTheCell()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var worksheet = workbook.AddWorksheet("Sheet1", workbook.CreateStyle(fill: new Fill(Color.Yellow)));

        var cell = worksheet.AddRow(workbook.CreateStyle(new Font { Bold = true })).AddCell();

        Assert.True(ResolvedFont(cell.Style).Bold?.Val?.Value);
        Assert.Equal("FFFFFF00", ForegroundColor(cell.Style));
    }

    [Fact]
    public void CellStyle_TakesPrecedenceOverRowStyleForTheSameFacet()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var worksheet = workbook.AddWorksheet("Sheet1");
        var row = worksheet.AddRow(workbook.CreateStyle(new Font { FontSize = 8 }));

        var cell = row.AddCell("value", workbook.CreateStyle(new Font { FontSize = 20 }));

        Assert.Equal(20d, ResolvedFont(cell.Style).FontSize?.Val?.Value);
    }

    [Fact]
    public void DifferentFacetsAtRowAndCellLevel_BothSurvive()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var worksheet = workbook.AddWorksheet("Sheet1");
        var row = worksheet.AddRow(workbook.CreateStyle(new Font { Bold = true }));

        var cell = row.AddCell("value", workbook.CreateStyle(fill: new Fill(Color.Yellow)));

        Assert.True(ResolvedFont(cell.Style).Bold?.Val?.Value);
        Assert.Equal("FFFFFF00", ForegroundColor(cell.Style));
    }

    [Fact]
    public void AllThreeLevels_ComposeWithTheNarrowestWinning()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var sheetStyle = workbook.CreateStyle(
            font: new Font { FontSize = 8 },
            border: new Border(BorderStyleValues.Thin));
        var worksheet = workbook.AddWorksheet("Sheet1", sheetStyle);
        var row = worksheet.AddRow(workbook.CreateStyle(fill: new Fill(Color.Yellow)));

        var cell = row.AddCell("value", workbook.CreateStyle(new Font { FontSize = 20 }));

        // Cell wins the font size, the row contributes the fill, the sheet contributes the border.
        Assert.Equal(20d, ResolvedFont(cell.Style).FontSize?.Val?.Value);
        Assert.Equal("FFFFFF00", ForegroundColor(cell.Style));
        OoxmlAssert.ChildOrder(ResolvedBorder(cell.Style), "left", "right", "top", "bottom");
    }

    [Fact]
    public void AddStyleOnAnExistingCell_LayersOnTopOfWhatItAlreadyHad()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var worksheet = workbook.AddWorksheet("Sheet1");
        var cell = worksheet.AddCell("value", workbook.CreateStyle(new Font { Bold = true }));

        cell.AddStyle(workbook.CreateStyle(fill: new Fill(Color.Yellow)));

        Assert.True(ResolvedFont(cell.Style).Bold?.Val?.Value);
        Assert.Equal("FFFFFF00", ForegroundColor(cell.Style));
    }

    [Fact]
    public void SheetStyleAppliedToLaterRows_StillReachesCellsAddedAfterwards()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var worksheet = workbook.AddWorksheet("Sheet1", workbook.CreateStyle(new Font { Italic = true }));

        worksheet.AddRow();
        var later = worksheet.AddRow().AddCell();

        Assert.True(ResolvedFont(later.Style).Italic?.Val?.Value);
    }

    [Fact]
    public void ResolvedFont_CarriesTheWorkbookDefaultsAlongsideTheRequestedFacet()
    {
        // Pinning the surprise explicitly: asking only for bold yields a font that also carries
        // the default size, colour and name, because the merge runs against the default entry.
        using var workbook = CreateInMemorySpreadsheet();
        var worksheet = workbook.AddWorksheet("Sheet1");

        var cell = worksheet.AddRow(workbook.CreateStyle(new Font { Bold = true })).AddCell();

        OoxmlAssert.RendersAs(
            """
            <font>
              <b val="1" /><sz val="11" /><color rgb="FF000000" /><name val="Calibri" />
            </font>
            """,
            ResolvedFont(cell.Style));
    }
}
