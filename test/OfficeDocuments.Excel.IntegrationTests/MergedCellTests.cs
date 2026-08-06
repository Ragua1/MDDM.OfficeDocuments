using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Excel.Interfaces;
using OfficeDocuments.Excel.Styles;
using OfficeDocuments.Excel.TestKit;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.IntegrationTests;

/// <summary>
/// What <c>AddCellOnRange</c> and <see cref="IRange.Merge"/> put into the file.
/// </summary>
/// <remarks>
/// Argument validation for the same methods lives with the type under test, in
/// <c>WorksheetTest</c> and <c>RowTest</c>; this file is about the produced <c>mergeCells</c>.
/// Until these tests existed the only merge reference asserted anywhere was the one
/// <c>RangeAndAdvancedFeaturesTest</c> checks for <see cref="IRange.Merge"/> — the ordinary
/// horizontal <c>AddCellOnRange</c>, which is the common case, was never looked at in the output.
/// </remarks>
public class MergedCellTests : SpreadsheetTestBase
{
    [Fact]
    public void AddCellOnRange_HorizontalRangeOnRow_WritesTheExpectedMergeReference()
    {
        using var stream = new MemoryStream();
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            sheet.AddRow().AddCellOnRange(2, 4);
        }

        Assert.Equal(["B1:D1"], MergeReferences(stream));
    }

    [Fact]
    public void AddCellOnRange_HorizontalRangeOnWorksheet_WritesTheExpectedMergeReference()
    {
        using var stream = new MemoryStream();
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            sheet.AddCellOnRange(2, 4, 3);
        }

        Assert.Equal(["B3:D3"], MergeReferences(stream));
    }

    [Fact]
    public void AddCellOnRange_Block_WritesTheExpectedMergeReference()
    {
        using var stream = new MemoryStream();
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            sheet.AddCellOnRange(2, 4, 2, 5);
        }

        Assert.Equal(["B2:D5"], MergeReferences(stream));
    }

    [Fact]
    public void AddCellOnRange_SingleColumnAcrossRows_WritesVerticalMerge()
    {
        using var stream = new MemoryStream();
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            sheet.AddCellOnRange(2, 2, 1, 3);
        }

        Assert.Equal(["B1:B3"], MergeReferences(stream));
    }

    [Fact]
    public void AddCellOnRange_SingleCellOnWorksheet_CreatesCellAndWritesNoMerge()
    {
        using var stream = new MemoryStream();
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet("Sheet 1");

            var cell = sheet.AddCellOnRange(2, 2, 3, 3);

            Assert.Equal((uint)2, cell.ColumnIndex);
            Assert.Equal((uint)3, cell.RowIndex);
        }

        // A one-cell range is not a merge, so no MergeCells element belongs in the file at all.
        Assert.Null(MergeCellsElement(stream));
    }

    [Fact]
    public void AddCellOnRange_SingleColumnOnRow_CreatesCellAndWritesNoMerge()
    {
        using var stream = new MemoryStream();
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet("Sheet 1");

            var cell = sheet.AddRow().AddCellOnRange(2, 2);

            Assert.Equal((uint)2, cell.ColumnIndex);
        }

        Assert.Null(MergeCellsElement(stream));
    }

    [Fact]
    public void AddCellOnRange_RowAndWorksheetOverloads_ProduceTheSameMerge()
    {
        using var stream = new MemoryStream();
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            sheet.AddRow(1).AddCellOnRange(2, 4);
            sheet.AddCellOnRange(2, 4, 2);
        }

        // The two entry points reach the merge by different code paths. They drifted apart once
        // already, on whether a single-column range was valid, because nothing compared them.
        var references = MergeReferences(stream);
        Assert.Equal(["B1:D1", "B2:D2"], references);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void AddCellOnRange_WithStyle_AppliesItToEveryCellOfTheRange(bool throughRow)
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var style = w.CreateStyle(new Font { FontSize = 20 });
        var row = sheet.AddRow(1);

        if (throughRow)
        {
            row.AddCellOnRange(2, 4, style);
        }
        else
        {
            sheet.AddCellOnRange(2, 4, 1, style);
        }

        // Only the top-left cell was ever asserted before, so a style that reached the returned
        // cell but not the rest of the range would have gone unnoticed.
        for (uint column = 2; column <= 4; column++)
        {
            var cell = sheet.GetCell(column, 1);
            Assert.NotNull(cell);
            Assert.NotNull(cell.Style);
            Assert.Equal(20d, StylesheetProbe.Font(cell.Style).FontSize?.Val?.Value);
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void AddCellOnRange_OnAStyledRow_KeepsTheCallerStyleOnEveryCell(bool throughRow)
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var row = sheet.AddRow(1, w.CreateStyle(new Font { Bold = true, FontSize = 9 }));

        if (throughRow)
        {
            row.AddCellOnRange(2, 4, w.CreateStyle(new Font { FontSize = 20 }));
        }
        else
        {
            sheet.AddCellOnRange(2, 4, 1, w.CreateStyle(new Font { FontSize = 20 }));
        }

        // The documented precedence is that the narrower level wins a facet both set. Applying the
        // row style again to a cell that already exists inverts it, and the worksheet path used to
        // do exactly that because Range.Merge touches every cell a second time after ApplyStyle.
        for (uint column = 2; column <= 4; column++)
        {
            var style = sheet.GetCell(column, 1)?.Style;
            Assert.NotNull(style);
            var font = StylesheetProbe.Font(style);
            Assert.Equal(20d, font.FontSize?.Val?.Value);
            Assert.True(font.Bold?.Val?.Value);
        }
    }

    [Fact]
    public void AddCellOnRange_SingleCellOnStyledRow_KeepsTheCallerStyle()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        sheet.AddRow(1, w.CreateStyle(new Font { Bold = true, FontSize = 9 }));

        var cell = sheet.AddCellOnRange(2, 2, 1, w.CreateStyle(new Font { FontSize = 20 }));

        // The single-cell path reaches the returned cell differently from the merging one, so it
        // needs its own check that the row style did not get stamped back over the caller's.
        Assert.NotNull(cell.Style);
        var font = StylesheetProbe.Font(cell.Style);
        Assert.Equal(20d, font.FontSize?.Val?.Value);
        Assert.True(font.Bold?.Val?.Value);
    }

    [Fact]
    public void AddCellOnRange_SameRangeTwice_WritesOneMerge()
    {
        using var stream = new MemoryStream();
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            sheet.AddCellOnRange(1, 4, 1);
            sheet.AddCellOnRange(1, 4, 1);
        }

        Assert.Equal(["A1:D1"], MergeReferences(stream));
    }

    [Fact]
    public void AddCellOnRange_OverlappingRange_ThrowsArgumentException()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        sheet.AddCellOnRange(1, 4, 1);

        var exception = Assert.Throws<ArgumentException>(() => sheet.AddCellOnRange(3, 6, 1));

        Assert.Contains("A1:D1", exception.Message);
    }

    [Fact]
    public void Merge_OverlappingRange_ThrowsArgumentException()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        sheet.GetRange("B2:D4").Merge();

        // Overlap on a single corner cell is still overlap; Excel reports the workbook as damaged.
        Assert.Throws<ArgumentException>(() => sheet.GetRange("D4:F6").Merge());
    }

    [Fact]
    public void Merge_TouchingButNotOverlappingRanges_BothSurvive()
    {
        using var stream = new MemoryStream();
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            sheet.GetRange("B2:D4").Merge();
            sheet.GetRange("E2:G4").Merge();
            sheet.GetRange("B5:D7").Merge();
        }

        Assert.Equal(["B2:D4", "E2:G4", "B5:D7"], MergeReferences(stream));
    }

    [Fact]
    public void Merge_FirstMergeRejected_LeavesNoEmptyMergeCellsElement()
    {
        using var stream = new MemoryStream();
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            sheet.GetRange("B2:D4").Merge();
            Assert.Throws<ArgumentException>(() => sheet.GetRange("C3:E5").Merge());
            sheet.GetRange("B2:D4").Merge();
        }

        // CT_MergeCells requires at least one child, so a rejected merge must not be able to
        // leave an empty <mergeCells/> behind — that alone would make the document schema-invalid.
        Assert.Equal(["B2:D4"], MergeReferences(stream));
    }

    private static SpreadsheetLib.MergeCells? MergeCellsElement(Stream workbook)
    {
        workbook.Position = 0;
        using var document = SpreadsheetDocument.Open(workbook, false);
        var worksheetPart = WorkbookParts.GetWorksheetPart(document, "Sheet 1");
        var worksheetElement = worksheetPart.Worksheet
                               ?? throw new InvalidOperationException("Worksheet element was not found.");

        return worksheetElement.GetFirstChild<SpreadsheetLib.MergeCells>();
    }

    private static string[] MergeReferences(Stream workbook) =>
        (MergeCellsElement(workbook) ?? throw new InvalidOperationException("MergeCells element was not found."))
        .Elements<SpreadsheetLib.MergeCell>()
        .Select(mergeCell => mergeCell.Reference?.Value ?? string.Empty)
        .ToArray();
}
