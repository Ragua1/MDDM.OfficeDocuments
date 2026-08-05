using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeDocuments.Excel.Extensions;
using OfficeDocuments.Excel.Options;
using OfficeDocuments.Excel.TestKit;
using Color = System.Drawing.Color;
using StyleBorder = OfficeDocuments.Excel.Styles.Border;
using StyleFill = OfficeDocuments.Excel.Styles.Fill;
using StyleFont = OfficeDocuments.Excel.Styles.Font;

namespace OfficeDocuments.Excel.IntegrationTests.Styles;

/// <summary>
/// Conditional formatting stores its formatting in <c>dxfs</c> (differential formats), a separate
/// collection from <c>cellXfs</c> with its own deduplication.
/// </summary>
/// <remarks>
/// This path had no coverage at all: rules were only ever checked for existence, never for what
/// they actually point at, so a rule referencing the wrong dxf — or a dxfs collection growing one
/// entry per rule — would have gone unnoticed.
/// </remarks>
public class DifferentialFormatTests : SpreadsheetTestBase
{
    private static DifferentialFormats ReadDifferentialFormats(Stream workbook)
    {
        workbook.Position = 0;
        using var document = SpreadsheetDocument.Open(workbook, false);
        var stylesheet = document.WorkbookPart?.WorkbookStylesPart?.Stylesheet
                         ?? throw new InvalidOperationException("The workbook has no stylesheet.");

        return stylesheet.DifferentialFormats ?? new DifferentialFormats();
    }

    private static List<uint> RuleFormatIds(Stream workbook, string worksheetName)
    {
        workbook.Position = 0;
        using var document = SpreadsheetDocument.Open(workbook, false);
        var worksheet = WorkbookParts.GetWorksheetPart(document, worksheetName).Worksheet;

        return worksheet.Elements<ConditionalFormatting>()
            .SelectMany(formatting => formatting.Elements<ConditionalFormattingRule>())
            .Select(rule => rule.FormatId?.Value ?? uint.MaxValue)
            .ToList();
    }

    [Fact]
    public void ConditionalFormatting_AllocatesADifferentialFormatCarryingTheStyle()
    {
        using var workbook = CreateInMemorySpreadsheet(out var stream);
        var worksheet = workbook.AddWorksheet("Sheet1");
        var style = workbook.CreateStyle(fill: new StyleFill(Color.Yellow));

        worksheet.GetRange("A1:A5").AddConditionalFormatting(ConditionalFormattingOptions.ContainsText("x", style));
        workbook.Close();

        var differentialFormats = ReadDifferentialFormats(stream);
        var format = differentialFormats.Elements<DifferentialFormat>().Single();

        Assert.Equal("FFFFFF00", format.Fill?.PatternFill?.ForegroundColor?.Rgb?.Value);
        Assert.Equal([0u], RuleFormatIds(stream, "Sheet1"));
    }

    [Fact]
    public void TwoRulesWithTheSameStyle_ShareOneDifferentialFormat()
    {
        using var workbook = CreateInMemorySpreadsheet(out var stream);
        var worksheet = workbook.AddWorksheet("Sheet1");
        var style = workbook.CreateStyle(fill: new StyleFill(Color.Yellow));

        worksheet.GetRange("A1:A5").AddConditionalFormatting(ConditionalFormattingOptions.ContainsText("x", style));
        worksheet.GetRange("B1:B5").AddConditionalFormatting(ConditionalFormattingOptions.ContainsText("y", style));
        workbook.Close();

        Assert.Single(ReadDifferentialFormats(stream).Elements<DifferentialFormat>());
        Assert.Equal([0u, 0u], RuleFormatIds(stream, "Sheet1"));
    }

    [Fact]
    public void TwoRulesWithDistinctStyles_GetDistinctDifferentialFormats()
    {
        using var workbook = CreateInMemorySpreadsheet(out var stream);
        var worksheet = workbook.AddWorksheet("Sheet1");
        var yellow = workbook.CreateStyle(fill: new StyleFill(Color.Yellow));
        var red = workbook.CreateStyle(fill: new StyleFill(Color.Red));

        worksheet.GetRange("A1:A5").AddConditionalFormatting(ConditionalFormattingOptions.ContainsText("x", yellow));
        worksheet.GetRange("B1:B5").AddConditionalFormatting(ConditionalFormattingOptions.ContainsText("y", red));
        workbook.Close();

        var formats = ReadDifferentialFormats(stream).Elements<DifferentialFormat>().ToList();

        Assert.Equal(2, formats.Count);
        Assert.Equal([0u, 1u], RuleFormatIds(stream, "Sheet1"));
        Assert.Equal("FFFFFF00", formats[0].Fill?.PatternFill?.ForegroundColor?.Rgb?.Value);
        Assert.Equal("FFFF0000", formats[1].Fill?.PatternFill?.ForegroundColor?.Rgb?.Value);
    }

    [Fact]
    public void DifferentialFormat_CarriesFontFillAndBorderButNotTheDefaults()
    {
        using var workbook = CreateInMemorySpreadsheet(out var stream);
        var worksheet = workbook.AddWorksheet("Sheet1");
        var style = workbook.CreateStyle(
            font: new StyleFont { Bold = true },
            fill: new StyleFill(Color.Yellow),
            border: new StyleBorder(Enums.BorderStyleValues.Thin));

        worksheet.GetRange("A1:A5").AddConditionalFormatting(ConditionalFormattingOptions.GreaterThan("10", style));
        workbook.Close();

        var format = ReadDifferentialFormats(stream).Elements<DifferentialFormat>().Single();

        Assert.True(format.Font?.Bold?.Val?.Value);
        Assert.Equal("FFFFFF00", format.Fill?.PatternFill?.ForegroundColor?.Rgb?.Value);
        Assert.NotNull(format.Border);
        Assert.Null(format.NumberingFormat);
    }

    [Fact]
    public void RulesAcrossTwoWorksheets_ShareTheWorkbookWideDifferentialFormats()
    {
        // dxfs live on the workbook stylesheet, not per worksheet.
        using var workbook = CreateInMemorySpreadsheet(out var stream);
        var first = workbook.AddWorksheet("First");
        var second = workbook.AddWorksheet("Second");
        var style = workbook.CreateStyle(fill: new StyleFill(Color.Yellow));

        first.GetRange("A1:A5").AddConditionalFormatting(ConditionalFormattingOptions.ContainsText("x", style));
        second.GetRange("A1:A5").AddConditionalFormatting(ConditionalFormattingOptions.ContainsText("x", style));
        workbook.Close();

        Assert.Single(ReadDifferentialFormats(stream).Elements<DifferentialFormat>());
        Assert.Equal([0u], RuleFormatIds(stream, "First"));
        Assert.Equal([0u], RuleFormatIds(stream, "Second"));
    }

    [Fact]
    public void TwoColorScale_NeedsNoDifferentialFormat()
    {
        using var workbook = CreateInMemorySpreadsheet(out var stream);
        var worksheet = workbook.AddWorksheet("Sheet1");

        worksheet.GetRange("A1:A5")
            .AddConditionalFormatting(ConditionalFormattingOptions.TwoColorScale(Color.White, Color.Green));
        workbook.Close();

        Assert.Empty(ReadDifferentialFormats(stream).Elements<DifferentialFormat>());
    }

    [Fact]
    public void ManyRulesReusingOneStyle_DoNotGrowTheDxfsCollection()
    {
        using var workbook = CreateInMemorySpreadsheet(out var stream);
        var worksheet = workbook.AddWorksheet("Sheet1");
        var style = workbook.CreateStyle(fill: new StyleFill(Color.Yellow));

        for (var column = 1; column <= 25; column++)
        {
            var reference = $"{CellExtension.GetExcelCellReference((uint)column, 1)}:{CellExtension.GetExcelCellReference((uint)column, 5)}";
            worksheet.GetRange(reference).AddConditionalFormatting(ConditionalFormattingOptions.ContainsText("x", style));
        }

        workbook.Close();

        Assert.Single(ReadDifferentialFormats(stream).Elements<DifferentialFormat>());
        Assert.All(RuleFormatIds(stream, "Sheet1"), formatId => Assert.Equal(0u, formatId));
    }
}
