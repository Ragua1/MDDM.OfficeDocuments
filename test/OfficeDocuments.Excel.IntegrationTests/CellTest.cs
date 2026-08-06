using OfficeDocuments.Excel.TestKit;
using System.Globalization;
using OfficeDocuments.Excel.Interfaces;
using Color = System.Drawing.Color;
using Styles_Font = OfficeDocuments.Excel.Styles.Font;
using Styles_NumberingFormat = OfficeDocuments.Excel.Styles.NumberingFormat;

namespace OfficeDocuments.Excel.IntegrationTests;

public class CellTest : SpreadsheetTestBase
{
    [Fact]
    public void CreateCell()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var cell = sheet.AddCell();
        Assert.NotNull(cell);
        Assert.IsAssignableFrom<ICell>(cell);

        Assert.NotNull(sheet.CurrentRow);
        Assert.Contains(cell, sheet.CurrentRow.Cells);
        Assert.Equal(sheet.CurrentRow.RowIndex, cell.RowIndex);
    }

    [Fact]
    public void CreateCellWithValue()
    {
        var filePath = GetFilepath("doc2.xlsx");
        using (var w = CreateNewSpreadsheet(filePath))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            const string value = "Aloha";
            var cell = sheet.AddCell(value);
            Assert.Equal(value, cell.Value);
        }
        Assert.True(File.Exists(filePath));
    }

    [Fact]
    public void SetString()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var value = "Aloha";
        var cell1 = sheet.AddCell(value);
        var cell2 = sheet.AddCell();
        cell2.Value = value;
        Assert.Equal(value, cell1.Value);
        Assert.Equal(value, cell2.Value);
        Assert.Equal(cell1.Value, cell2.Value);

        Assert.NotNull(cell1.Style);
        Assert.Equal(49, cell1.Style.NumberFormatId);
    }

    [Fact]
    public void SetInteger()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var value = 165752313;
        var cell1 = sheet.AddCell(value);
        var cell2 = sheet.AddCell();
        cell2.SetValue(value);
        Assert.Equal(
            value.ToString(),
            cell1.Value
        );
        Assert.Equal(
            value.ToString(),
            cell2.Value
        );
        Assert.Equal(cell1.Value, cell2.Value);

        Assert.NotNull(cell1.Style);
        Assert.Equal(1, cell1.Style.NumberFormatId);
    }

    [Fact]
    public void SetIntegerWithStyle()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var s = w.CreateStyle(numberFormat: new Styles_NumberingFormat("#,##0x"));
        var value = 98435123;
        var cell1 = sheet.AddCell(value, s);
        var cell2 = sheet.AddCell(s);
        cell2.SetValue(value);
        Assert.Equal(
            value.ToString(),
            cell1.Value
        );
        Assert.Equal(
            value.ToString(),
            cell2.Value
        );
        Assert.Equal(cell1.Value, cell2.Value);

        var excelUserFormatsIndex = 170;
        Assert.NotNull(cell1.Style);
        Assert.Equal(excelUserFormatsIndex, cell1.Style.NumberFormatId);
    }

    [Fact]
    public void SetDouble()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var value = 165752313.216546;
        var cell1 = sheet.AddCell(value);
        var cell2 = sheet.AddCell();
        cell2.SetValue(value);
        Assert.Equal(
            value.ToString(CultureInfo.InvariantCulture),
            cell1.Value
        );
        Assert.Equal(
            value.ToString(CultureInfo.InvariantCulture),
            cell2.Value
        );
        Assert.Equal(cell1.Value, cell2.Value);
    }

    [Fact]
    public void SetDoubleWithStyle()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var s = w.CreateStyle(numberFormat: new Styles_NumberingFormat("#,##0.##x"));
        var value = 645.541;
        var cell1 = sheet.AddCell(value, s);
        var cell2 = sheet.AddCell(s);
        cell2.SetValue(value);
        Assert.Equal(
            value.ToString(CultureInfo.InvariantCulture),
            cell1.Value
        );
        Assert.Equal(
            value.ToString(CultureInfo.InvariantCulture),
            cell2.Value
        );
        Assert.Equal(cell1.Value, cell2.Value);

        var excelUserFormatsIndex = 170;
        Assert.NotNull(cell1.Style);
        Assert.Equal(excelUserFormatsIndex, cell1.Style.NumberFormatId);
    }

    [Fact]
    public void SetLong()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var value = 165752313216546;
        var cell1 = sheet.AddCell(value);
        var cell2 = sheet.AddCell();
        cell2.SetValue(value);
        Assert.Equal(
            value.ToString(CultureInfo.InvariantCulture),
            cell1.Value
        );
        Assert.Equal(
            value.ToString(CultureInfo.InvariantCulture),
            cell2.Value
        );
        Assert.Equal(cell1.Value, cell2.Value);
    }

    [Fact]
    public void SetLongWithStyle()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var s = w.CreateStyle(numberFormat: new Styles_NumberingFormat("#,##0x"));
        var value = 165752313216546;
        var cell1 = sheet.AddCell(value, s);
        var cell2 = sheet.AddCell(s);
        cell2.SetValue(value);
        Assert.Equal(
            value.ToString(CultureInfo.InvariantCulture),
            cell1.Value
        );
        Assert.Equal(
            value.ToString(CultureInfo.InvariantCulture),
            cell2.Value
        );
        Assert.Equal(cell1.Value, cell2.Value);

        var excelUserFormatsIndex = 170;
        Assert.NotNull(cell1.Style);
        Assert.Equal(excelUserFormatsIndex, cell1.Style.NumberFormatId);
    }

    [Fact]
    public void SetDecimal()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var value = 165752313216546.6516511m;
        var cell1 = sheet.AddCell(value);
        var cell2 = sheet.AddCell();
        cell2.SetValue(value);
        Assert.Equal(
            value.ToString(CultureInfo.InvariantCulture),
            cell1.Value
        );
        Assert.Equal(
            value.ToString(CultureInfo.InvariantCulture),
            cell2.Value
        );
        Assert.Equal(cell1.Value, cell2.Value);
    }

    [Fact]
    public void SetDecimalWithStyle()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var s = w.CreateStyle(numberFormat: new Styles_NumberingFormat("#,##0.##0x"));
        var value = 16575231321654.6565465426m;
        var cell1 = sheet.AddCell(value, s);
        var cell2 = sheet.AddCell(s);
        cell2.SetValue(value);
        Assert.Equal(
            value.ToString(CultureInfo.InvariantCulture),
            cell1.Value
        );
        Assert.Equal(
            value.ToString(CultureInfo.InvariantCulture),
            cell2.Value
        );
        Assert.Equal(cell1.Value, cell2.Value);

        var excelUserFormatsIndex = 170;
        Assert.NotNull(cell1.Style);
        Assert.Equal(excelUserFormatsIndex, cell1.Style.NumberFormatId);
    }

    [Fact]
    public void SetDate()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var value = DateTime.Now;
        var cell1 = sheet.AddCell(value);
        var cell2 = sheet.AddCell();
        cell2.SetValue(value);
        Assert.Equal(
            value.ToOADate().ToString(CultureInfo.InvariantCulture),
            cell1.Value
        );
        Assert.Equal(
            value.ToOADate().ToString(CultureInfo.InvariantCulture),
            cell2.Value
        );
        Assert.Equal(cell1.Value, cell2.Value);

        Assert.NotNull(cell1.Style);
        Assert.Equal(14, cell1.Style.NumberFormatId);
    }

    [Fact]
    public void SetDateWithStyle()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var s = w.CreateStyle(numberFormat: new Styles_NumberingFormat("d/m/yyyy H:mm:ss"));
        var value = DateTime.Now;
        var cell1 = sheet.AddCell(value, s);
        var cell2 = sheet.AddCell(s);
        cell2.SetValue(value);
        Assert.Equal(
            value.ToOADate().ToString(CultureInfo.InvariantCulture),
            cell1.Value
        );
        Assert.Equal(
            value.ToOADate().ToString(CultureInfo.InvariantCulture),
            cell2.Value
        );
        Assert.Equal(cell1.Value, cell2.Value);

        var excelUserFormatsIndex = 170;
        Assert.NotNull(cell1.Style);
        Assert.Equal(excelUserFormatsIndex, cell1.Style.NumberFormatId);
    }

    [Fact]
    public void SetBoolean()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var value = true;
        var cell1 = sheet.AddCell(value);
        var cell2 = sheet.AddCell();
        cell2.SetValue(value);
        Assert.Equal("1", cell1.Value);
        Assert.Equal("1", cell2.Value);
        Assert.Equal(cell1.Value, cell2.Value);
        Assert.True(cell1.GetBoolValue());
        Assert.True(cell2.GetBoolValue());
    }

    [Fact]
    public void SetFormula()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var formula = "SUM(A1:A5)";
        var cell = sheet.AddCell();
        cell.SetFormula(formula);
        Assert.Equal(formula, cell.GetFormula());
    }

    [Fact]
    public void SumInRangeFormula()
    {
        int[] numbers = new int[] { 1, 2, 3 };
        int sum = 0;
        foreach (int num in numbers)
        {
            sum += num;
        }

        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var formula = "SUM(B1:D1)";

        var cell1 = sheet.AddCellWithFormula(formula);
        var cell2 = sheet.AddCell(numbers[0]);
        var cell3 = sheet.AddCell(numbers[1]);
        var cell4 = sheet.AddCell(numbers[2]);

        double value = cell1.GetFormulaValue();

        Assert.Equal(sum, value);
    }

    [Fact]
    public void SumInRangeFormula2()
    {
        int[] numbers = new int[] { 1, 2, 3 };
        int sum = 0;
        foreach (int num in numbers)
        {
            sum += num;
        }

        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var formula = "SUM(B1:H1)";

        var cell1 = sheet.AddCellWithFormula(formula);
        var cell2 = sheet.AddCell(numbers[0]);
        var cell3 = sheet.AddCell(numbers[1]);
        var cell4 = sheet.AddCell(numbers[2]);

        double value = cell1.GetFormulaValue();

        Assert.Equal(sum, value);
    }

    [Fact]
    public void SumInRangeFormula3()
    {
        int[] numbers = new int[] { 1, 2, 3 };
        int sum = 0;
        foreach (int num in numbers)
        {
            sum += num;
        }

        string text = "Lorem Ipsum";

        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var formula = "SUM(B1:D1)";

        var cell1 = sheet.AddCellWithFormula(formula);
        var cell2 = sheet.AddCell(numbers[0]);
        var cell3 = sheet.AddCell(numbers[1]);
        var cell4 = sheet.AddCell(text);

        Assert.Throws<ArgumentException>(() => cell1.GetFormulaValue());
    }

    [Fact]
    public void SumInRangeFormula4()
    {
        int[] numbers = new int[] { 1, 2, 3, 4 };
        int sum = 0;
        foreach (int num in numbers)
        {
            sum += num;
        }

        string formula2 = "SUM(E1:F1)";

        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var formula = "SUM(B1:D1)";

        var cell1 = sheet.AddCellWithFormula(formula);
        var cell2 = sheet.AddCell(numbers[0]);
        var cell3 = sheet.AddCell(numbers[1]);
        var cell4 = sheet.AddCell(formula2);
        var cell5 = sheet.AddCell(numbers[2]);
        var cell6 = sheet.AddCell(numbers[3]);

        Assert.Throws<ArgumentException>(() => cell1.GetFormulaValue());
    }

    [Fact]
    public void SumInRangeFormula5()
    {
        int[] numbers = new int[] { 1, 2, 3, 4 };
        int sum = 0;
        foreach (int num in numbers)
        {
            sum += num;
        }

        string formula2 = "SUM(E1:F1)";

        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var formula = "SUM(B1:D1)";

        var cell1 = sheet.AddCellWithFormula(formula);
        var cell2 = sheet.AddCell(numbers[0]);
        var cell3 = sheet.AddCell(numbers[1]);
        var cell4 = sheet.AddCellWithFormula(formula2);
        var cell5 = sheet.AddCell(numbers[2]);
        var cell6 = sheet.AddCell(numbers[3]);

        double value = cell1.GetFormulaValue();

        Assert.Equal(sum, value);
    }

    [Fact]
    public void CellInheritStyleFromRow()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var s = w.CreateStyle(new Styles_Font { Color = Color.DarkGoldenrod });
        var row = sheet.AddRow(s);
        var cell = row.AddCell();

        Assert.NotNull(row.Style);
        Assert.NotNull(cell.Style);
        Assert.Equal(row.Style.StyleIndex, cell.Style.StyleIndex);
    }

    [Fact]
    public void CellInheritStyleFromSheet()
    {
        using var w = CreateInMemorySpreadsheet();
        var s = w.CreateStyle(new Styles_Font { Color = Color.DarkGoldenrod });
        var sheet = w.AddWorksheet("Sheet 1", s);
        var cell = sheet.AddCell();

        Assert.NotNull(sheet.Style);
        Assert.NotNull(cell.Style);
        Assert.Equal(sheet.Style.StyleIndex, cell.Style.StyleIndex);
    }

    [Fact]
    public void CellHasCorrectIndexes1()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var cell = sheet.AddCellOnIndex(5, 3);

        Assert.Equal((uint)5, cell.ColumnIndex);
        Assert.Equal((uint)3, cell.RowIndex);
        Assert.Equal("E3", cell.CellReference);
    }

    [Fact]
    public void CellHasCorrectIndexes2()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        var row = sheet.AddRow(3);
        row.AddCell();
        row.AddCell();
        row.AddCell();
        row.AddCell();
        var cell = row.AddCell();

        Assert.Equal((uint)5, cell.ColumnIndex);
        Assert.Equal((uint)3, cell.RowIndex);
        Assert.Equal("E3", cell.CellReference);
    }

    [Fact]
    public void CellHasCorrectIndexes3()
    {
        using var stream = new MemoryStream();
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            var cell = sheet.AddCellOnIndex(5, 3);

            Assert.Equal((uint)5, cell.ColumnIndex);
            Assert.Equal((uint)3, cell.RowIndex);
            Assert.Equal("E3", cell.CellReference);
        }

        using (var w = OpenExistingSpreadsheet(stream))
        {
            var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
            Assert.NotNull(sheet);

            var row = sheet.GetRow(3);
            Assert.NotNull(row);

            var cell = row.GetCell(5);
            Assert.NotNull(cell);

            cell = sheet.GetCell(5, 3);
            Assert.NotNull(cell);

            Assert.Equal((uint)5, cell.ColumnIndex);
            Assert.Equal((uint)3, cell.RowIndex);
            Assert.Equal("E3", cell.CellReference);
        }
    }

    [Fact]
    public void CellSetAndGetBoolValue()
    {
        using var stream = new MemoryStream();
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet();
            var cell = sheet.AddCellOnIndex(5, 3);
            cell.SetValue(true);

            Assert.True(cell.GetBoolValue());
        }

        using (var w = OpenExistingSpreadsheet(stream))
        {
            var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
            Assert.NotNull(sheet);

            var cell = sheet.GetCell(5, 3);
            Assert.NotNull(cell);
            Assert.True(cell.GetBoolValue());

            bool res;
            Assert.True(cell.TryGetValue(out res));
            Assert.True(res);
        }
    }

    [Fact]
    public void CellSetAndGetIntValue()
    {
        using var stream = new MemoryStream();
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet();
            var cell = sheet.AddCellOnIndex(5, 3);
            cell.SetValue(int.MaxValue);

            Assert.Equal(int.MaxValue, cell.GetIntValue());
        }

        using (var w = OpenExistingSpreadsheet(stream))
        {
            var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
            Assert.NotNull(sheet);

            var cell = sheet.GetCell(5, 3);
            Assert.NotNull(cell);
            Assert.Equal(int.MaxValue, cell.GetIntValue());

            int res;
            Assert.True(cell.TryGetValue(out res));
            Assert.Equal(int.MaxValue, res);
        }
    }

    [Fact]
    public void CellSetAndGetLongValue()
    {
        using var stream = new MemoryStream();
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet();
            var cell = sheet.AddCellOnIndex(5, 3);
            cell.SetValue(long.MaxValue);

            Assert.Equal(long.MaxValue, cell.GetLongValue());
        }

        using (var w = OpenExistingSpreadsheet(stream))
        {
            var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
            Assert.NotNull(sheet);

            var cell = sheet.GetCell(5, 3);
            Assert.NotNull(cell);
            Assert.Equal(long.MaxValue, cell.GetLongValue());

            long res;
            Assert.True(cell.TryGetValue(out res));
            Assert.Equal(long.MaxValue, res);
        }
    }

    [Fact]
    public void CellSetAndGetDoubleValue()
    {
        using var stream = new MemoryStream();
        var value = 16831231.1564d;
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet();
            var cell = sheet.AddCellOnIndex(5, 3);
            cell.SetValue(value);

            Assert.Equal(value, cell.GetDoubleValue());
        }

        using (var w = OpenExistingSpreadsheet(stream))
        {
            var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
            Assert.NotNull(sheet);

            var cell = sheet.GetCell(5, 3);
            Assert.NotNull(cell);
            Assert.Equal(value, cell.GetDoubleValue());

            double res;
            Assert.True(cell.TryGetValue(out res));
            Assert.Equal(value, res);
        }
    }

    [Fact]
    public void CellSetAndGetDecimalValue()
    {
        using var stream = new MemoryStream();
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet();
            var cell = sheet.AddCellOnIndex(5, 3);
            cell.SetValue(decimal.MaxValue);

            Assert.Equal(decimal.MaxValue, cell.GetDecimalValue());
        }

        using (var w = OpenExistingSpreadsheet(stream))
        {
            var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
            Assert.NotNull(sheet);

            var cell = sheet.GetCell(5, 3);
            Assert.NotNull(cell);
            Assert.Equal(decimal.MaxValue, cell.GetDecimalValue());

            decimal res;
            Assert.True(cell.TryGetValue(out res));
            Assert.Equal(decimal.MaxValue, res);
        }
    }

    [Fact]
    public void CellSetAndGetStringValue()
    {
        using var stream = new MemoryStream();
        var value = "Alohomora";
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet();
            var cell = sheet.AddCellOnIndex(5, 3);
            cell.SetValue(value);

            Assert.Equal(value, cell.GetStringValue());
        }

        using (var w = OpenExistingSpreadsheet(stream))
        {
            var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
            Assert.NotNull(sheet);

            var cell = sheet.GetCell(5, 3);
            Assert.NotNull(cell);
            Assert.Equal(value, cell.GetStringValue());
        }
    }

    [Fact]
    public void CellSetAndGetDateValue()
    {
        using var stream = new MemoryStream();
        var value = DateTime.Now;
        var format = "dd.MM.yyyy hh:mm:ss";
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet();
            var cell = sheet.AddCellOnIndex(5, 3);
            cell.SetValue(value);

            Assert.Equal(value.ToString(format), cell.GetDateValue().ToString(format));
        }

        using (var w = OpenExistingSpreadsheet(stream))
        {
            var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
            Assert.NotNull(sheet);

            var cell = sheet.GetCell(5, 3);
            Assert.NotNull(cell);
            Assert.Equal(value.ToString(format), cell.GetDateValue().ToString(format));

            DateTime res;
            Assert.True(cell.TryGetValue(out res));
            Assert.Equal(value.ToString(format), res.ToString(format));
        }
    }

    [Fact]
    public void CellHasValue()
    {
        using var stream = new MemoryStream();
        var value = "Aika";
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet();
            var cell = sheet.AddCellOnIndex(5, 3);
            cell.SetValue(value);

            Assert.Equal(value, cell.GetStringValue());
        }

        using (var w = OpenExistingSpreadsheet(stream))
        {
            var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
            Assert.NotNull(sheet);

            var cell = sheet.GetCell(5, 3);
            Assert.NotNull(cell);
            Assert.True(cell.HasValue());
            Assert.False(cell.HasFormula());
        }
    }

    [Fact]
    public void CellHasFormula()
    {
        using var stream = new MemoryStream();
        var formula = "SUM(C1:C4)";
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet();
            var row = sheet.AddRow(3);
            row.AddCell(15);
            row.AddCell(4);
            row.AddCell(-2);
            row.AddCell(9);

            var cell = row.AddCellWithFormula(5, formula);
            Assert.Equal(formula, cell.GetFormula());
        }

        using (var w = OpenExistingSpreadsheet(stream))
        {
            var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
            Assert.NotNull(sheet);

            var cell = sheet.GetCell(5, 3);
            Assert.NotNull(cell);
            Assert.False(cell.HasValue());
            Assert.True(cell.HasFormula());
        }
    }

    [Fact]
    public void CellHasFormulaGetValue()
    {
        using var stream = new MemoryStream();
        var formula = "SUM(C1:C4)";
        using (var w = CreateNewSpreadsheet(stream))
        {
            var sheet = w.AddWorksheet();
            var row = sheet.AddRow(3);
            row.AddCell(15);
            row.AddCell(4);
            row.AddCell(-2);
            row.AddCell(9);

            var cell = row.AddCellWithFormula(5, formula);
            Assert.Equal(formula, cell.GetFormula());
        }

        using (var w = OpenExistingSpreadsheet(stream))
        {
            var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
            Assert.NotNull(sheet);

            var cell = sheet.GetCell(5, 3);
            Assert.NotNull(cell);
            Assert.False(cell.HasValue());
            Assert.True(cell.HasFormula());

            Assert.Throws<InvalidOperationException>(() => cell.GetIntValue());
        }
    }

    [Fact]
    public void CreateMultipleCellsWithValue()
    {
        var filePath = GetFilepath("doc27.xlsx");
        using (var w = CreateNewSpreadsheet(filePath))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            for (int i = 0; i < 10; i++)
            {
                var cell = sheet.AddCell(i);
                Assert.Equal(i, cell.GetIntValue());
            }
        }
        Assert.True(File.Exists(filePath));
    }

    [Fact]
    public void CreateMultipleCellsWithStyle()
    {
        var filePath = GetFilepath("doc28.xlsx");
        using (var w = CreateNewSpreadsheet(filePath))
        {
            var sheet = w.AddWorksheet("Sheet 1");
            for (int i = 0; i < 10; i++)
            {
                var s = w.CreateStyle(numberFormat: new Styles_NumberingFormat("#,##0x"));
                var value = i;
                var cell1 = sheet.AddCell(value, s);
                var cell2 = sheet.AddCell(s);
                cell2.SetValue(value);

                Assert.Equal(value, cell1.GetIntValue());
                Assert.Equal(value, cell2.GetIntValue());
                Assert.Equal(cell1.GetIntValue(), cell2.GetIntValue());
            }
        }
        Assert.True(File.Exists(filePath));
    }

    [Fact]
    public void CreateRandomCellsInRange()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");
        for (int i = 0; i < 10; i++)
        {
            sheet.AddRow();
            for (int j = 0; j < 10; j++)
            {
                Random rnd = new Random();
                int createCell = rnd.Next(0, 2);

                if (createCell == 0)
                {
                    var cell = sheet.AddCell(j);
                    Assert.Equal(j, cell.GetIntValue());
                }
                else
                {
                    var cell = sheet.AddCell();
                    Assert.NotNull(cell);
                }
            }
        }
    }

    [Fact]
    public void CountCellsWithValueInRange()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");

        var formula = "COUNT(B1:G1)";

        var cell1 = sheet.AddCellWithFormula(formula);

        var cell2 = sheet.AddCell(1);
        var cell3 = sheet.AddCell(2);
        var cell4 = sheet.AddCell(3);

        var cell5 = sheet.AddCell();
        var cell6 = sheet.AddCell();
        var cell7 = sheet.AddCell();

        var sum = cell1.GetFormulaValue();

        Assert.Equal(3, sum);
    }

    [Fact]
    public void CountCellsByArgument()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");

        var formula = "COUNTIF(B1:G1,\"car\")";

        var cell1 = sheet.AddCellWithFormula(formula);

        var cell2 = sheet.AddCell("car");
        var cell3 = sheet.AddCell("bike");
        var cell4 = sheet.AddCell("car");
        var cell5 = sheet.AddCell("train");
        var cell6 = sheet.AddCell("car");
        var cell7 = sheet.AddCell("plane");

        var sum = cell1.GetFormulaValue();

        Assert.Equal(3, sum);
    }

    [Fact]
    public void CountCellsByArgument2()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");

        var formula = "COUNTIF(B1:G1,B1)";

        var cell1 = sheet.AddCellWithFormula(formula);

        var cell2 = sheet.AddCell("car");
        var cell3 = sheet.AddCell("bike");
        var cell4 = sheet.AddCell("car");
        var cell5 = sheet.AddCell("train");
        var cell6 = sheet.AddCell("car");
        var cell7 = sheet.AddCell("plane");

        var sum = cell1.GetFormulaValue();

        Assert.Equal(3, sum);
    }

    [Fact]
    public void CountCellsByArgument3()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");

        var formula = "COUNTIF(B1:G1,B1)";

        var cell1 = sheet.AddCellWithFormula(formula);

        var cell2 = sheet.AddCell();
        var cell3 = sheet.AddCell("bike");
        var cell4 = sheet.AddCell();
        var cell5 = sheet.AddCell("train");
        var cell6 = sheet.AddCell();
        var cell7 = sheet.AddCell("plane");

        var sum = cell1.GetFormulaValue();

        Assert.Equal(3, sum);
    }

    [Fact]
    public void GetMedian()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");

        var formula = "MEDIAN(B1:F1)";

        var cell1 = sheet.AddCellWithFormula(formula);

        var cell2 = sheet.AddCell(1);
        var cell3 = sheet.AddCell(5);
        var cell4 = sheet.AddCell(7);
        var cell5 = sheet.AddCell(9);
        var cell6 = sheet.AddCell(2);

        var median = cell1.GetFormulaValue();
        var expected = 5;

        Assert.Equal(expected, median);
    }

    [Fact]
    public void GetMedian2()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");

        var formula = "MEDIAN(B1:G1)";

        var cell1 = sheet.AddCellWithFormula(formula);

        var cell2 = sheet.AddCell(1);
        var cell3 = sheet.AddCell(5);
        var cell4 = sheet.AddCell(7);
        var cell5 = sheet.AddCell(9);
        var cell6 = sheet.AddCell(2);
        var cell7 = sheet.AddCell(10);

        var median = cell1.GetFormulaValue();
        var expected = 6;

        Assert.Equal(expected, median);
    }

    [Fact]
    public void GetMedian3()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");

        var formula = "MEDIAN(B1:G1)";

        var cell1 = sheet.AddCellWithFormula(formula);

        var cell2 = sheet.AddCell(1);
        var cell3 = sheet.AddCell(5);
        var cell4 = sheet.AddCell("Lorem ipsum");
        var cell5 = sheet.AddCell(9);
        var cell6 = sheet.AddCell(2);
        var cell7 = sheet.AddCell("Dolor sit amet");

        var median = cell1.GetFormulaValue();
        var expected = 3.5;

        Assert.Equal(expected, median);
    }

    [Fact]
    public void SumInVerticalRange_ReturnsSum()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");

        var cell1 = sheet.AddCellWithFormula("SUM(B2:B4)");
        sheet.AddCell(2u, 2u, 10);
        sheet.AddCell(2u, 3u, 20);
        sheet.AddCell(2u, 4u, 30);

        Assert.Equal(60d, cell1.GetFormulaValue());
    }

    [Fact]
    public void MedianOfEmptyRange_Throws()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");

        var cell1 = sheet.AddCellWithFormula("MEDIAN(B1:D1)");

        Assert.Throws<InvalidOperationException>(() => cell1.GetFormulaValue());
    }

    [Fact]
    public void UnsupportedFunction_Throws()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");

        var cell1 = sheet.AddCellWithFormula("SUMIF(B1:D1,\">0\")");

        Assert.Throws<NotSupportedException>(() => cell1.GetFormulaValue());
    }

    [Fact]
    public void SumOfDecimalValues_ReturnsDouble()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");

        var cell1 = sheet.AddCellWithFormula("SUM(B1:D1)");
        sheet.AddCell(1.5);
        sheet.AddCell(2.25);
        sheet.AddCell(0.25);

        Assert.Equal(4.0d, cell1.GetFormulaValue());
    }

    [Fact]
    public void GetFormulaValue_WithoutFormula_Throws()
    {
        using var w = CreateInMemorySpreadsheet();
        var sheet = w.AddWorksheet("Sheet 1");

        var cell = sheet.AddCell(5);

        Assert.Throws<InvalidOperationException>(() => cell.GetFormulaValue());
    }
}