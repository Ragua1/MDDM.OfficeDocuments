using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Excel.TestKit;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Interfaces;
using OfficeDocuments.Excel.Styles;
using Color = System.Drawing.Color;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.IntegrationTests;

public class WorksheetTest : SpreadsheetTestBase
{
    [Fact]
    public void CreateCellOnWrongColumnIndex()
    {
            
        var exception = Assert.Throws<ArgumentException>(() =>
        {
            using var w = CreateInMemorySpreadsheet();
            var sheet = w.AddWorksheet("Sheet 1");
            sheet.AddCell(0, 0);
        });
    }

    [Fact]
    public void CreateCellOnWrongRowIndex()
    {
            
        var exception = Assert.Throws<ArgumentException>(() =>
        {
            using var w = CreateInMemorySpreadsheet();
            var sheet = w.AddWorksheet("Sheet 1");
            sheet.AddCell(5, 0, 0);
        });
    }

    [Fact]
    public void CreateCellOnSpecificColumnIndex()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var cell = sheet.AddCell(5);
        Assert.NotNull(cell);
        Assert.IsAssignableFrom<ICell>(cell);
    }

    [Fact]
    public void CreateCellOnSpecificRowIndexAndColumnIndex()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var cell = sheet.AddCell(5, 4);
        Assert.NotNull(cell);
        Assert.IsAssignableFrom<ICell>(cell);
    }

    [Fact]
    public void CreateCellWithStyle()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var s = w.CreateStyle(new Font { Color = Color.Blue }, new Fill(Color.BurlyWood), new Border(BorderStyleValues.Hair), new NumberingFormat("0"));
        var cell = sheet.AddCell(s);
        Assert.NotNull(cell.Style);
        Assert.True(cell.Style.FontId > 0);
        Assert.True(cell.Style.FillId > 0);
        Assert.True(cell.Style.BorderId > 0);
        Assert.True(cell.Style.NumberFormatId > 0);
        Assert.True(cell.Style.StyleIndex > 0);
    }

    [Fact]
    public void GetRow_SparseRows_ReturnsExpectedRows()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");

        sheet.AddRow(10).AddCell("Tenth");
        sheet.AddRow(3).AddCell("Third");
        sheet.AddRow(7).AddCell("Seventh");

        Assert.Equal("Third", sheet.GetRow(3)?.GetCell(1)?.GetStringValue());
        Assert.Equal("Seventh", sheet.GetRow(7)?.GetCell(1)?.GetStringValue());
        Assert.Equal("Tenth", sheet.GetRow(10)?.GetCell(1)?.GetStringValue());
    }

    [Fact]
    public void GetCellByReference_SparseWorksheet_ReturnsExpectedCell()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");

        sheet.AddCell(2, 3, "B3 value");
        sheet.AddCell(27, 10, "AA10 value");

        Assert.Equal("B3 value", sheet.GetCellByReference("b3")?.GetStringValue());
        Assert.Equal("AA10 value", sheet.GetCellByReference("AA10")?.GetStringValue());
    }

    [Fact]
    public void GetRange_InvalidReference_ThrowsArgumentException()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");

        Assert.Throws<ArgumentException>(() => sheet.GetRange("A0:B2"));
    }

    // The three overloads used to disagree: two rejected an inverted range by returning null, and
    // each had its own copy of the bounds check, so a single-column range was invalid on one path
    // and valid on another.
    [Theory]
    [InlineData(0u, 4u, "beginColumn")]
    [InlineData(5u, 4u, "endColumn")]
    public void AddCellOnRange_InvalidColumnRange_ThrowsArgumentException(uint beginColumn, uint endColumn, string paramName)
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");

        var onCurrentRow = Assert.Throws<ArgumentException>(() => sheet.AddCellOnRange(beginColumn, endColumn));
        var onRowIndex = Assert.Throws<ArgumentException>(() => sheet.AddCellOnRange(beginColumn, endColumn, 1));
        var onBlock = Assert.Throws<ArgumentException>(() => sheet.AddCellOnRange(beginColumn, endColumn, 1, 2));

        Assert.Equal(paramName, onCurrentRow.ParamName);
        Assert.Equal(paramName, onRowIndex.ParamName);
        Assert.Equal(paramName, onBlock.ParamName);
    }

    [Theory]
    [InlineData(0u, 2u, "beginRow")]
    [InlineData(3u, 2u, "endRow")]
    public void AddCellOnRange_InvalidRowRange_ThrowsArgumentException(uint beginRow, uint endRow, string paramName)
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");

        var exception = Assert.Throws<ArgumentException>(() => sheet.AddCellOnRange(2, 4, beginRow, endRow));

        Assert.Equal(paramName, exception.ParamName);
    }

    [Fact]
    public void AddCellOnRange_RowIndexBelowOne_ThrowsArgumentException()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");

        var exception = Assert.Throws<ArgumentException>(() => sheet.AddCellOnRange(2, 4, 0));

        Assert.Equal("rowIndex", exception.ParamName);
    }

    [Fact]
    public void AddCellOnRange_SingleCell_CreatesCellAndWritesNoMerge()
    {
        using var stream = new MemoryStream();
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet("Sheet 1");

            var cell = sheet.AddCellOnRange(2, 2, 3, 3);

            Assert.Equal((uint)2, cell.ColumnIndex);
            Assert.Equal((uint)3, cell.RowIndex);
        }

        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = WorkbookParts.GetWorksheetPart(document, "Sheet 1");
        var worksheetElement = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet element was not found.");
        // A one-cell range is not a merge, so no MergeCells element belongs in the file at all.
        Assert.Null(worksheetElement.GetFirstChild<SpreadsheetLib.MergeCells>());
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

        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = WorkbookParts.GetWorksheetPart(document, "Sheet 1");
        var worksheetElement = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet element was not found.");
        var mergeCells = worksheetElement.GetFirstChild<SpreadsheetLib.MergeCells>() ?? throw new InvalidOperationException("MergeCells element was not found.");
        var mergeCell = Assert.Single(mergeCells.Elements<SpreadsheetLib.MergeCell>());

        Assert.Equal("B1:B3", mergeCell.Reference?.Value);
    }
}