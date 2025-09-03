using OfficeDocuments.Excel.Interfaces;
using OfficeDocuments.Excel.Styles;
using Color = System.Drawing.Color;

namespace OfficeDocuments.Excel.Tests;

public class RowTest : SpreadsheetTestBase
{
    [Fact]
    public void CreateRow()
    {
        var filePath = GetFilepath("doc1.xlsx");
        using (var w = CreateTestee(filePath))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            var row = sheet.AddRow();
            Assert.NotNull(row);
            Assert.NotNull(row.Element);
            Assert.IsAssignableFrom<IRow>(row);

            Assert.Contains(row, sheet.Rows);
            Assert.Equal(sheet.CurrentRow.RowIndex, row.RowIndex);
            Assert.Null(row.CurrentCell);
        }
    }

    [Fact]
    public void CreateRowWithStyle()
    {
        var filePath = GetFilepath("doc2.xlsx");
        using (var w = CreateTestee(filePath))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            var s = w.CreateStyle(new Font { Color = Color.Coral }, new Fill(Color.Black));
            var row = sheet.AddRow(s);

            Assert.True(row.Style.StyleIndex > 0);
        }
    }

    [Fact]
    public void CreateRowOnSpecificRowIndex()
    {
        var filePath = GetFilepath("doc3.xlsx");
        using (var w = CreateTestee(filePath))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            var row = sheet.AddRow(5);

            Assert.Equal((uint)5, row.RowIndex);
        }
    }

    [Fact]
    public void CreateRowOnWrongRowIndex()
    {
        var filePath = GetFilepath("doc4.xlsx");
            
        var exception = Assert.Throws<ArgumentException>(() =>
        {
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                sheet.AddRow(0);
            }
        });
    }

    [Fact]
    public void CreateRowAndAddCell()
    {
        var filePath = GetFilepath("doc5.xlsx");
        using (var w = CreateTestee(filePath))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            var row = sheet.AddRow();
            var cell1 = row.AddCell();
            var cell2 = row.AddCellOnIndex(3);

            Assert.Equal((uint)1, cell1.ColumnIndex);
            Assert.Equal((uint)3, cell2.ColumnIndex);
        }
    }

    [Fact]
    public void CreateRowAndAddCellWithValue()
    {
        var filePath = GetFilepath("doc6.xlsx");
        using (var w = CreateTestee(filePath))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            var value = "Alea iacta est";
            var row = sheet.AddRow();
            var cell1 = row.AddCell(value);
            var cell2 = row.AddCell(3, value);

            Assert.Equal((uint)1, cell1.ColumnIndex);
            Assert.Equal(value, cell1.Value);

            Assert.Equal((uint)3, cell2.ColumnIndex);
            Assert.Equal(value, cell2.Value);

            value = "Sumilian Eri Lopte";
            cell1 = row.AddCell(1, value);
            Assert.Equal((uint)1, cell1.ColumnIndex);
            Assert.Equal(value, cell1.Value);
        }
    }

    [Fact]
    public void CreateRowAndAddCellWithFormula()
    {
        var filePath = GetFilepath("doc7.xlsx");
        using (var w = CreateTestee(filePath))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            var value = "Sum(A1:A2)";
            var row = sheet.AddRow();
            var cell = row.AddCellWithFormula(value);
            var cell2 = row.AddCellWithFormula(5, value);

            Assert.Equal((uint)1, cell.ColumnIndex);
            Assert.Equal(value, cell.Element.CellFormula.Text);

            value = "Sum(B1:B2)";
            cell = row.AddCellWithFormula(1, value);
            Assert.Equal((uint)1, cell.ColumnIndex);
            Assert.Equal(value, cell.Element.CellFormula.Text);
        }
    }

    [Fact]
    public void CreateRowWithValueOnWrongRowIndex()
    {
        var filePath = GetFilepath("doc8.xlsx");
            
        var exception = Assert.Throws<ArgumentException>(() =>
        {
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                sheet.AddCell(0, 0);
            }
        });
    }

    [Fact]
    public void CreateRowWithFormulaOnWrongRowIndex()
    {
        var filePath = GetFilepath("doc9.xlsx");
            
        var exception = Assert.Throws<ArgumentException>(() =>
        {
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                sheet.AddCellWithFormula(0, "0");
            }
        });
    }

    [Fact]
    public void CreateRowAndCellOnRange()
    {
        var filePath = GetFilepath("doc10.xlsx");
        using (var w = CreateTestee(filePath))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            var row = sheet.AddRow();
            row.AddCellOnRange(2, 4);

            Assert.NotNull(row.GetCell(2));
            Assert.NotNull(row.GetCell(3));
            Assert.NotNull(row.GetCell(4));
        }
    }

    [Fact]
    public void CreateRowAndCellOnRangeOnWrongColumnIndex()
    {
        var filePath = GetFilepath("doc11.xlsx");
            
        var exception = Assert.Throws<ArgumentException>(() =>
        {
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var row = sheet.AddRow();
                row.AddCellOnRange(0, 4);
            }
        });
    }

    [Fact]
    public void CreateRowAndCellOnWrongRange()
    {
        var filePath = GetFilepath("doc12.xlsx");
        using (var w = CreateTestee(filePath))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            var row = sheet.AddRow();
            var cell = row.AddCellOnRange(5, 4);

            Assert.Null(cell);
        }
    }

    [Fact]
    public void CreateRowAndCellOnBigRange()
    {
        var filePath = GetFilepath("doc13.xlsx");
        using (var w = CreateTestee(filePath))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            var row = sheet.AddRow();
            var cell = row.AddCellOnRange(2, 611);

            Assert.NotNull(cell);
        }
    }

    [Fact]
    public void GetCellByName()
    {
        var filePath = GetFilepath("doc14.xlsx");
        using (var w = CreateTestee(filePath))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            var value = "Alea iacta est";
            var row = sheet.AddRow();
            row.AddCell(value);
            row.AddCell(3, value);

            var cell1 = row.GetCell("A");
            Assert.NotNull(cell1);
            Assert.Equal((uint)1, cell1.ColumnIndex);
            Assert.Equal(value, cell1.Value);

            var cell2 = row.GetCell("C");
            Assert.NotNull(cell2);
            Assert.Equal((uint)3, cell2.ColumnIndex);
            Assert.Equal(value, cell2.Value);

            value = "Sumilian Eri Lopte";
            row.AddCell(1, value);

            cell1 = row.GetCell("A");
            Assert.Equal((uint)1, cell1.ColumnIndex);
            Assert.Equal(value, cell1.Value);
        }
    }
}