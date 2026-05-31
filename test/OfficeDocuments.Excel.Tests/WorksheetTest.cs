using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Interfaces;
using OfficeDocuments.Excel.Styles;
using Color = System.Drawing.Color;

namespace OfficeDocuments.Excel.Tests;

public class WorksheetTest : SpreadsheetTestBase
{
    [Fact]
    public void CreateCellOnWrongColumnIndex()
    {
        var filePath = GetFilepath("doc1.xlsx");
            
        var exception = Assert.Throws<ArgumentException>(() =>
        {
            using var w = CreateNewSpreadsheet(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            sheet.AddCell(0, 0);
        });
    }

    [Fact]
    public void CreateCellOnWrongRowIndex()
    {
        var filePath = GetFilepath("doc2.xlsx");
            
        var exception = Assert.Throws<ArgumentException>(() =>
        {
            using var w = CreateNewSpreadsheet(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            sheet.AddCell(5, 0, 0);
        });
    }

    [Fact]
    public void CreateCellOnSpecificColumnIndex()
    {
        var filePath = GetFilepath("doc3.xlsx");
        using var w = CreateNewSpreadsheet(filePath);
        var sheet = w.AddWorksheet("Sheet 1");
        var cell = sheet.AddCell(5);
        Assert.NotNull(cell);
        Assert.NotNull(cell.Element);
        Assert.IsAssignableFrom<ICell>(cell);
    }

    [Fact]
    public void CreateCellOnSpecificRowIndexAndColumnIndex()
    {
        var filePath = GetFilepath("doc4.xlsx");
        using var w = CreateNewSpreadsheet(filePath);
        var sheet = w.AddWorksheet("Sheet 1");
        var cell = sheet.AddCell(5, 4);
        Assert.NotNull(cell);
        Assert.NotNull(cell.Element);
        Assert.IsAssignableFrom<ICell>(cell);
    }

    [Fact]
    public void CreateCellWithStyle()
    {
        var filePath = GetFilepath("doc5.xlsx");
        using var w = CreateNewSpreadsheet(filePath);
        var sheet = w.AddWorksheet("Sheet 1");
        var s = w.CreateStyle(new Font { Color = Color.Blue }, new Fill(Color.BurlyWood), new Border(BorderStyleValues.Hair), new NumberingFormat("0"));
        var cell = sheet.AddCell(s);
        Assert.True(cell.Style.FontId > 0);
        Assert.True(cell.Style.FillId > 0);
        Assert.True(cell.Style.BorderId > 0);
        Assert.True(cell.Style.NumberFormatId > 0);
        Assert.True(cell.Style.StyleIndex > 0);
    }

    [Fact]
    public void GetRow_SparseRows_ReturnsExpectedRows()
    {
        var filePath = GetFilepath("doc6.xlsx");
        using var w = CreateNewSpreadsheet(filePath);
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
        var filePath = GetFilepath("doc7.xlsx");
        using var w = CreateNewSpreadsheet(filePath);
        var sheet = w.AddWorksheet("Sheet 1");

        sheet.AddCell(2, 3, "B3 value");
        sheet.AddCell(27, 10, "AA10 value");

        Assert.Equal("B3 value", sheet.GetCellByReference("b3")?.GetStringValue());
        Assert.Equal("AA10 value", sheet.GetCellByReference("AA10")?.GetStringValue());
    }

    [Fact]
    public void GetRange_InvalidReference_ThrowsArgumentException()
    {
        var filePath = GetFilepath("doc8.xlsx");
        using var w = CreateNewSpreadsheet(filePath);
        var sheet = w.AddWorksheet("Sheet 1");

        Assert.Throws<ArgumentException>(() => sheet.GetRange("A0:B2"));
    }
}