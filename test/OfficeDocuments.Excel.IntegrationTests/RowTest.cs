using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Excel.TestKit;
using OfficeDocuments.Excel.Interfaces;
using OfficeDocuments.Excel.Styles;
using Color = System.Drawing.Color;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.IntegrationTests;

public class RowTest : SpreadsheetTestBase
{
    [Fact]
    public void CreateRow()
    {
        using (var w = CreateInMemorySpreadsheet())
        {
            var sheet = w.AddWorksheet("Sheet 1");
            var row = sheet.AddRow();
            Assert.NotNull(row);
            Assert.IsAssignableFrom<IRow>(row);

            Assert.Contains(row, sheet.Rows);
            Assert.NotNull(sheet.CurrentRow);
            Assert.Equal(sheet.CurrentRow.RowIndex, row.RowIndex);
            Assert.Null(row.CurrentCell);
        }
    }

    [Fact]
    public void CreateRowWithStyle()
    {
        using (var w = CreateInMemorySpreadsheet())
        {
            var sheet = w.AddWorksheet("Sheet 1");
            var s = w.CreateStyle(new Font { Color = Color.Coral }, new Fill(Color.Black));
            var row = sheet.AddRow(s);

            Assert.NotNull(row.Style);
            Assert.True(row.Style.StyleIndex > 0);
        }
    }

    [Fact]
    public void CreateRowOnSpecificRowIndex()
    {
        using (var w = CreateInMemorySpreadsheet())
        {
            var sheet = w.AddWorksheet("Sheet 1");
            var row = sheet.AddRow(5);

            Assert.Equal((uint)5, row.RowIndex);
        }
    }

    [Fact]
    public void CreateRowOnWrongRowIndex()
    {
            
        var exception = Assert.Throws<ArgumentException>(() =>
        {
            using (var w = CreateInMemorySpreadsheet())
            {
                var sheet = w.AddWorksheet("Sheet 1");
                sheet.AddRow(0);
            }
        });
    }

    [Fact]
    public void CreateRowAndAddCell()
    {
        using (var w = CreateInMemorySpreadsheet())
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
        using (var w = CreateInMemorySpreadsheet())
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
        using (var w = CreateInMemorySpreadsheet())
        {
            var sheet = w.AddWorksheet("Sheet 1");
            var value = "Sum(A1:A2)";
            var row = sheet.AddRow();
            var cell = row.AddCellWithFormula(value);
            var cell2 = row.AddCellWithFormula(5, value);

            Assert.Equal((uint)1, cell.ColumnIndex);
            Assert.Equal(value, cell.GetFormula());

            value = "Sum(B1:B2)";
            cell = row.AddCellWithFormula(1, value);
            Assert.Equal((uint)1, cell.ColumnIndex);
            Assert.Equal(value, cell.GetFormula());
        }
    }

    [Fact]
    public void CreateRowWithValueOnWrongRowIndex()
    {
            
        var exception = Assert.Throws<ArgumentException>(() =>
        {
            using (var w = CreateInMemorySpreadsheet())
            {
                var sheet = w.AddWorksheet("Sheet 1");
                sheet.AddCell(0, 0);
            }
        });
    }

    [Fact]
    public void CreateRowWithFormulaOnWrongRowIndex()
    {
            
        var exception = Assert.Throws<ArgumentException>(() =>
        {
            using (var w = CreateInMemorySpreadsheet())
            {
                var sheet = w.AddWorksheet("Sheet 1");
                sheet.AddCellWithFormula(0, "0");
            }
        });
    }

    [Fact]
    public void CreateRowAndCellOnRange()
    {
        using (var w = CreateInMemorySpreadsheet())
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
            
        var exception = Assert.Throws<ArgumentException>(() =>
        {
            using (var w = CreateInMemorySpreadsheet())
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var row = sheet.AddRow();
                row.AddCellOnRange(0, 4);
            }
        });
    }

    [Fact]
    public void AddCellOnRange_EndColumnBeforeBeginColumn_ThrowsArgumentException()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var row = sheet.AddRow();

        var exception = Assert.Throws<ArgumentException>(() => row.AddCellOnRange(5, 4));

        Assert.Equal("endColumn", exception.ParamName);
    }

    [Fact]
    public void AddCellOnRange_SingleColumn_CreatesCellAndWritesNoMerge()
    {
        using var stream = new MemoryStream();
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            var row = sheet.AddRow();

            var cell = row.AddCellOnRange(2, 2);

            Assert.Equal((uint)2, cell.ColumnIndex);
        }

        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = WorkbookParts.GetWorksheetPart(document, "Sheet 1");
        var worksheetElement = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet element was not found.");
        // A one-cell range is not a merge, so no MergeCells element belongs in the file at all.
        Assert.Null(worksheetElement.GetFirstChild<SpreadsheetLib.MergeCells>());
    }

    [Fact]
    public void CreateRowAndCellOnBigRange()
    {
        using (var w = CreateInMemorySpreadsheet())
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
        using (var w = CreateInMemorySpreadsheet())
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
            Assert.NotNull(cell1);
            Assert.Equal((uint)1, cell1.ColumnIndex);
            Assert.Equal(value, cell1.Value);
        }
    }

    [Fact]
    public void GetCellByMultiLetterColumnName_ReturnsExpectedCell()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var row = sheet.AddRow();
        var value = "AA value";

        row.AddCell(27, value);

        var cell = row.GetCell("AA");

        Assert.NotNull(cell);
        Assert.Equal((uint)27, cell.ColumnIndex);
        Assert.Equal(value, cell.Value);
    }
}
