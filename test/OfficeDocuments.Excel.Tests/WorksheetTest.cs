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
            using var w = CreateTestee(filePath);
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
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            sheet.AddCell(5, 0, 0);
        });
    }

    [Fact]
    public void CreateCellOnSpecificColumnIndex()
    {
        var filePath = GetFilepath("doc3.xlsx");
        using var w = CreateTestee(filePath);
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
        using var w = CreateTestee(filePath);
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
        using var w = CreateTestee(filePath);
        var sheet = w.AddWorksheet("Sheet 1");
        var s = w.CreateStyle(new Font { Color = Color.Blue }, new Fill(Color.BurlyWood), new Border(BorderStyleValues.Hair), new NumberingFormat("0"));
        var cell = sheet.AddCell(s);
        Assert.True(cell.Style.FontId > 0);
        Assert.True(cell.Style.FillId > 0);
        Assert.True(cell.Style.BorderId > 0);
        Assert.True(cell.Style.NumberFormatId > 0);
        Assert.True(cell.Style.StyleIndex > 0);
    }
}