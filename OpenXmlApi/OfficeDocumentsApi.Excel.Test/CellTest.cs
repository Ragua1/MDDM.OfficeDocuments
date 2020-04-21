using System;
using System.Globalization;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDocumentsApi.Excel.Interfaces;
using Color = System.Drawing.Color;
using Font = OfficeDocumentsApi.Excel.Styles.Font;
using NumberingFormat = OfficeDocumentsApi.Excel.Styles.NumberingFormat;

namespace OfficeDocumentsApi.Excel.Test
{
    [TestClass]
    public class CellTest : SpreadsheetTestBase
    {
        [TestMethod]
        public void CreateCell()
        {
            var filePath = GetFilepath("doc1.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var cell = sheet.AddCell();
                Assert.IsNotNull(cell, "New cell cannot be null");
                Assert.IsNotNull(cell.Element, "Cells element cannot be null");
                Assert.IsInstanceOfType(cell, typeof(ICell), "Expected ICell type.");

                Assert.IsTrue(sheet.CurrentRow.Cells.Contains(cell));
                Assert.IsTrue(sheet.CurrentRow.RowIndex == cell.RowIndex);
            }
        }

        [TestMethod]
        public void CreateCellWithValue()
        {
            var filePath = GetFilepath("doc2.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                const string value = "Aloha";
                var cell = sheet.AddCellWithValue(value);
                Assert.AreEqual(cell.Value, value);
            }
        }

        [TestMethod]
        public void SetString()
        {
            var filePath = GetFilepath("doc3.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var value = "Aloha";
                var cell1 = sheet.AddCellWithValue(value);
                var cell2 = sheet.AddCell();
                cell2.Value = value;
                Assert.AreEqual(value, cell1.Value, $"Cell value of 'cell.SetValue()' is '{cell1.Value}', expected {value}");
                Assert.AreEqual(value, cell2.Value, $"Cell value of 'cell.SetValue()' is '{cell2.Value}', expected {value}");
                Assert.AreEqual((object) cell1.Value, cell2.Value, $"Cell values are not same");

                Assert.AreEqual(cell1.Style.NumberFormatId, 49, $"Number format is '{cell1.Style.NumberFormatId}', expected 49-'@'");
            }
        }

        [TestMethod]
        public void SetInteger()
        {
            var filePath = GetFilepath("doc4.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var value = 165752313;
                var cell1 = sheet.AddCellWithValue(value);
                var cell2 = sheet.AddCell();
                cell2.SetValue(value);
                Assert.AreEqual(
                    value.ToString(),
                    cell1.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell1.Value}', expected {value}"
                );
                Assert.AreEqual(
                    value.ToString(),
                    cell2.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell2.Value}', expected {value}"
                );
                Assert.AreEqual((object) cell1.Value, cell2.Value, $"Cell values are not same");

                Assert.AreEqual(cell1.Style.NumberFormatId, 1, $"Number format is '{cell1.Style.NumberFormatId}', expected 1-'0'");
            }
        }

        [TestMethod]
        public void SetIntegerWithStyle()
        {
            var filePath = GetFilepath("doc5.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var s = w.CreateStyle(numberFormat: new NumberingFormat("#,##0x"));
                var value = 98435123;
                var cell1 = sheet.AddCellWithValue(value, s);
                var cell2 = sheet.AddCell(s);
                cell2.SetValue(value);
                Assert.AreEqual(
                    value.ToString(),
                    cell1.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell1.Value}', expected {value}"
                );
                Assert.AreEqual(
                    value.ToString(),
                    cell2.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell2.Value}', expected {value}"
                );
                Assert.AreEqual((object) cell1.Value, cell2.Value, $"Cell values are not same");

                var excelUserFormatsIndex = 170;
                Assert.IsTrue(cell1.Style.NumberFormatId >= excelUserFormatsIndex, $"Number format is '{cell1.Style.NumberFormatId}', expected {excelUserFormatsIndex}+");
            }
        }

        [TestMethod]
        public void SetDouble()
        {
            var filePath = GetFilepath("doc6.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var value = 165752313.216546;
                var cell1 = sheet.AddCellWithValue(value);
                var cell2 = sheet.AddCell();
                cell2.SetValue(value);
                Assert.AreEqual(
                    value.ToString(CultureInfo.InvariantCulture),
                    cell1.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell1.Value}', expected {value}"
                );
                Assert.AreEqual(
                    value.ToString(CultureInfo.InvariantCulture),
                    cell2.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell2.Value}', expected {value}"
                );
                Assert.AreEqual((object) cell1.Value, cell2.Value, $"Cell values are not same");
            }
        }

        [TestMethod]
        public void SetDoubleWithStyle()
        {
            var filePath = GetFilepath("doc7.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var s = w.CreateStyle(numberFormat: new NumberingFormat("#,##0.##x"));
                var value = 645.541;
                var cell1 = sheet.AddCellWithValue(value, s);
                var cell2 = sheet.AddCell(s);
                cell2.SetValue(value);
                Assert.AreEqual(
                    value.ToString(CultureInfo.InvariantCulture),
                    cell1.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell1.Value}', expected {value}"
                );
                Assert.AreEqual(
                    value.ToString(CultureInfo.InvariantCulture),
                    cell2.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell2.Value}', expected {value}"
                );
                Assert.AreEqual((object) cell1.Value, cell2.Value, $"Cell values are not same");

                var excelUserFormatsIndex = 170;
                Assert.IsTrue(cell1.Style.NumberFormatId > excelUserFormatsIndex, $"Number format is '{cell1.Style.NumberFormatId}', expected {excelUserFormatsIndex}+");
            }
        }

        [TestMethod]
        public void SetLong()
        {
            var filePath = GetFilepath("doc100.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var value = 165752313216546;
                var cell1 = sheet.AddCellWithValue(value);
                var cell2 = sheet.AddCell();
                cell2.SetValue(value);
                Assert.AreEqual(
                    value.ToString(CultureInfo.InvariantCulture),
                    cell1.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell1.Value}', expected {value}"
                );
                Assert.AreEqual(
                    value.ToString(CultureInfo.InvariantCulture),
                    cell2.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell2.Value}', expected {value}"
                );
                Assert.AreEqual((object) cell1.Value, cell2.Value, $"Cell values are not same");
            }
        }

        [TestMethod]
        public void SetLongWithStyle()
        {
            var filePath = GetFilepath("doc101.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var s = w.CreateStyle(numberFormat: new NumberingFormat("#,##0x"));
                var value = 165752313216546;
                var cell1 = sheet.AddCellWithValue(value, s);
                var cell2 = sheet.AddCell(s);
                cell2.SetValue(value);
                Assert.AreEqual(
                    value.ToString(CultureInfo.InvariantCulture),
                    cell1.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell1.Value}', expected {value}"
                );
                Assert.AreEqual(
                    value.ToString(CultureInfo.InvariantCulture),
                    cell2.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell2.Value}', expected {value}"
                );
                Assert.AreEqual((object) cell1.Value, cell2.Value, $"Cell values are not same");

                var excelUserFormatsIndex = 170;
                Assert.IsTrue(cell1.Style.NumberFormatId >= excelUserFormatsIndex, $"Number format is '{cell1.Style.NumberFormatId}', expected {excelUserFormatsIndex}+");
            }
        }

        [TestMethod]
        public void SetDecimal()
        {
            var filePath = GetFilepath("doc102.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var value = 165752313216546.6516511m;
                var cell1 = sheet.AddCellWithValue(value);
                var cell2 = sheet.AddCell();
                cell2.SetValue(value);
                Assert.AreEqual(
                    value.ToString(CultureInfo.InvariantCulture),
                    cell1.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell1.Value}', expected {value.ToString(CultureInfo.InvariantCulture)}"
                );
                Assert.AreEqual(
                    value.ToString(CultureInfo.InvariantCulture),
                    cell2.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell2.Value}', expected {value.ToString(CultureInfo.InvariantCulture)}"
                );
                Assert.AreEqual((object) cell1.Value, cell2.Value, "Cell values are not same");
            }
        }

        [TestMethod]
        public void SetDecimalWithStyle()
        {
            var filePath = GetFilepath("doc103.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var s = w.CreateStyle(numberFormat: new NumberingFormat("#,##0.##0x"));
                var value = 16575231321654.6565465426m;
                var cell1 = sheet.AddCellWithValue(value, s);
                var cell2 = sheet.AddCell(s);
                cell2.SetValue(value);
                Assert.AreEqual(
                    value.ToString(CultureInfo.InvariantCulture),
                    cell1.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell1.Value}', expected {value.ToString(CultureInfo.InvariantCulture)}"
                );
                Assert.AreEqual(
                    value.ToString(CultureInfo.InvariantCulture),
                    cell2.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell2.Value}', expected {value.ToString(CultureInfo.InvariantCulture)}"
                );
                Assert.AreEqual((object) cell1.Value, cell2.Value, $"Cell values are not same");

                var excelUserFormatsIndex = 170;
                Assert.IsTrue(
                    cell1.Style.NumberFormatId >= excelUserFormatsIndex, $"Number format is '{cell1.Style.NumberFormatId}', expected {excelUserFormatsIndex}+"
                );
            }
        }

        [TestMethod]
        public void SetDate()
        {
            var filePath = GetFilepath("doc8.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var value = DateTime.Now;
                var cell1 = sheet.AddCellWithValue(value);
                var cell2 = sheet.AddCell();
                cell2.SetValue(value);
                Assert.AreEqual(
                    value.ToOADate().ToString(CultureInfo.InvariantCulture),
                    cell1.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell1.Value}', expected {value.ToOADate()}");
                Assert.AreEqual(
                    value.ToOADate().ToString(CultureInfo.InvariantCulture),
                    cell2.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell2.Value}', expected {value.ToOADate()}");
                Assert.AreEqual((object) cell1.Value, cell2.Value, $"Cell values are not same");

                Assert.AreEqual(cell1.Style.NumberFormatId, 14, $"Number format is '{cell1.Style.NumberFormatId}', expected 14-'d/m/yyyy'");
            }
        }

        [TestMethod]
        public void SetDateWithStyle()
        {
            var filePath = GetFilepath("doc9.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var s = w.CreateStyle(numberFormat: new NumberingFormat("d/m/yyyy H:mm:ss"));
                var value = DateTime.Now;
                var cell1 = sheet.AddCellWithValue(value, s);
                var cell2 = sheet.AddCell(s);
                cell2.SetValue(value);
                Assert.AreEqual(
                    value.ToOADate().ToString(CultureInfo.InvariantCulture),
                    cell1.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell1.Value}', expected {value.ToOADate()}"
                );
                Assert.AreEqual(
                    value.ToOADate().ToString(CultureInfo.InvariantCulture),
                    cell2.Value,
                    $"Cell value of 'cell.SetValue()' is '{cell2.Value}', expected {value.ToOADate()}"
                );
                Assert.AreEqual((object) cell1.Value, cell2.Value, $"Cell values are not same");

                var excelUserFormatsIndex = 170;
                Assert.IsTrue(
                    cell1.Style.NumberFormatId >= excelUserFormatsIndex,
                    $"Number format is '{cell1.Style.NumberFormatId}', expected '{excelUserFormatsIndex}'+"
                );
            }
        }

        [TestMethod]
        public void SetBoolean()
        {
            var filePath = GetFilepath("doc10.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var value = true;
                var cell1 = sheet.AddCellWithValue(value);
                var cell2 = sheet.AddCell();
                cell2.SetValue(value);
                Assert.AreEqual(cell1.Value, value.ToString(), $"Cell value of 'cell.SetValue()' is '{cell1.Value}', expected {value}");
                Assert.AreEqual(cell2.Value, value.ToString(), $"Cell value of 'cell.SetValue()' is '{cell2.Value}', expected {value}");
                Assert.AreEqual((object) cell1.Value, cell2.Value, $"Cell values are not same");
            }
        }

        [TestMethod]
        public void SetFormula()
        {
            var filePath = GetFilepath("doc11.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var formula = "Sum(A1:A5)";
                var cell = sheet.AddCell();
                cell.SetFormula(formula);
                Assert.AreEqual(cell.Element.CellFormula.Text, formula, $"Cell firmula is '{cell.Element.CellFormula.Text}', expected {formula}");
            }
        }

        [TestMethod]
        public void CellInheritStyleFromRow()
        {
            var filePath = GetFilepath("doc12.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var s = w.CreateStyle(new Font { Color = Color.DarkGoldenrod });
                var row = sheet.AddRow(s);
                var cell = row.AddCell();

                Assert.AreEqual(
                    (object) row.Style.StyleIndex, cell.Style.StyleIndex,
                    $"Cell not inherit style. Row style '{row.Style.StyleIndex}', cell stylw {cell.Style.StyleIndex}"
                );
            }
        }

        [TestMethod]
        public void CellInheritStyleFromSheet()
        {
            var filePath = GetFilepath("doc13.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var s = w.CreateStyle(new Font { Color = Color.DarkGoldenrod });
                var sheet = w.AddWorksheet("Sheet 1", s);
                var cell = sheet.AddCell();

                Assert.AreEqual(
                    (object) sheet.Style.StyleIndex, cell.Style.StyleIndex,
                    $"Cell not inherit style. Sheet style '{sheet.Style.StyleIndex}', cell style {cell.Style.StyleIndex}"
                );
            }
        }

        [TestMethod]
        public void CellHasCorrectIndexes1()
        {
            var filePath = GetFilepath("doc14.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var cell = sheet.AddCell(5, 3);

                Assert.AreEqual((uint)5, cell.ColumnIndex);
                Assert.AreEqual((uint)3, cell.RowIndex);
                Assert.AreEqual("E3", cell.CellReference);
            }
        }

        [TestMethod]
        public void CellHasCorrectIndexes2()
        {
            var filePath = GetFilepath("doc15.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var row = sheet.AddRow(3);
                row.AddCell();
                row.AddCell();
                row.AddCell();
                row.AddCell();
                var cell = row.AddCell();

                Assert.AreEqual((uint)5, cell.ColumnIndex);
                Assert.AreEqual((uint)3, cell.RowIndex);
                Assert.AreEqual("E3", cell.CellReference);
            }
        }

        [TestMethod]
        public void CellHasCorrectIndexes3()
        {
            var filePath = GetFilepath("doc16.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var cell = sheet.AddCell(5, 3);

                Assert.AreEqual((uint)5, cell.ColumnIndex);
                Assert.AreEqual((uint)3, cell.RowIndex);
                Assert.AreEqual("E3", cell.CellReference);
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.Worksheets.FirstOrDefault();
                Assert.IsNotNull(sheet);

                var row = sheet.GetRow(3);
                Assert.IsNotNull(row);

                var cell = row.GetCell(5);
                Assert.IsNotNull(cell);

                cell = sheet.GetCell(5, 3);
                Assert.IsNotNull(cell);

                Assert.AreEqual((uint)5, cell.ColumnIndex);
                Assert.AreEqual((uint)3, cell.RowIndex);
                Assert.AreEqual("E3", cell.CellReference);
            }
        }

        [TestMethod]
        public void CellSetAndGetBoolValue()
        {
            var filePath = GetFilepath("doc17.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet();
                var cell = sheet.AddCell(5, 3);
                cell.SetValue(true);

                Assert.AreEqual(true, cell.GetBoolValue());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.Worksheets.FirstOrDefault();
                Assert.IsNotNull(sheet);

                var cell = sheet.GetCell(5, 3);
                Assert.IsNotNull(cell);
                Assert.AreEqual(true, cell.GetBoolValue());

                bool res;
                Assert.AreEqual(true, cell.TryGetValue(out res));
                Assert.AreEqual(true, res);
            }
        }

        [TestMethod]
        public void CellSetAndGetIntValue()
        {
            var filePath = GetFilepath("doc18.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet();
                var cell = sheet.AddCell(5, 3);
                cell.SetValue(int.MaxValue);

                Assert.AreEqual(int.MaxValue, cell.GetIntValue());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.Worksheets.FirstOrDefault();
                Assert.IsNotNull(sheet);

                var cell = sheet.GetCell(5, 3);
                Assert.IsNotNull(cell);
                Assert.AreEqual(int.MaxValue, cell.GetIntValue());

                int res;
                Assert.AreEqual(true, cell.TryGetValue(out res));
                Assert.AreEqual(int.MaxValue, res);
            }
        }

        [TestMethod]
        public void CellSetAndGetLongValue()
        {
            var filePath = GetFilepath("doc19.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet();
                var cell = sheet.AddCell(5, 3);
                cell.SetValue(long.MaxValue);

                Assert.AreEqual(long.MaxValue, cell.GetLongValue());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.Worksheets.FirstOrDefault();
                Assert.IsNotNull(sheet);

                var cell = sheet.GetCell(5, 3);
                Assert.IsNotNull(cell);
                Assert.AreEqual(long.MaxValue, cell.GetLongValue());

                long res;
                Assert.AreEqual(true, cell.TryGetValue(out res));
                Assert.AreEqual(long.MaxValue, res);
            }
        }

        [TestMethod]
        public void CellSetAndGetDoubleValue()
        {
            var filePath = GetFilepath("doc19.xlsx");
            var value = 16831231.1564d;
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet();
                var cell = sheet.AddCell(5, 3);
                cell.SetValue(value);

                Assert.AreEqual(value, cell.GetDoubleValue());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.Worksheets.FirstOrDefault();
                Assert.IsNotNull(sheet);

                var cell = sheet.GetCell(5, 3);
                Assert.IsNotNull(cell);
                Assert.AreEqual(value, cell.GetDoubleValue());

                double res;
                Assert.AreEqual(true, cell.TryGetValue(out res));
                Assert.AreEqual(value, res);
            }
        }

        [TestMethod]
        public void CellSetAndGetDecimalValue()
        {
            var filePath = GetFilepath("doc20.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet();
                var cell = sheet.AddCell(5, 3);
                cell.SetValue(decimal.MaxValue);

                Assert.AreEqual(decimal.MaxValue, cell.GetDecimalValue());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.Worksheets.FirstOrDefault();
                Assert.IsNotNull(sheet);

                var cell = sheet.GetCell(5, 3);
                Assert.IsNotNull(cell);
                Assert.AreEqual(decimal.MaxValue, cell.GetDecimalValue());

                decimal res;
                Assert.AreEqual(true, cell.TryGetValue(out res));
                Assert.AreEqual(decimal.MaxValue, res);
            }
        }

        [TestMethod]
        public void CellSetAndGetStringValue()
        {
            var filePath = GetFilepath("doc21.xlsx");
            var value = "Alohomora";
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet();
                var cell = sheet.AddCell(5, 3);
                cell.SetValue(value);

                Assert.AreEqual(value, cell.GetStringValue());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.Worksheets.FirstOrDefault();
                Assert.IsNotNull(sheet);

                var cell = sheet.GetCell(5, 3);
                Assert.IsNotNull(cell);
                Assert.AreEqual(value, cell.GetStringValue());
            }
        }

        [TestMethod]
        public void CellSetAndGetDateValue()
        {
            var filePath = GetFilepath("doc22.xlsx");
            var value = DateTime.Now;
            var format = "dd.MM.yyyy hh:mm:ss";
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet();
                var cell = sheet.AddCell(5, 3);
                cell.SetValue(value);

                Assert.AreEqual(value.ToString(format), cell.GetDateValue().ToString(format));
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.Worksheets.FirstOrDefault();
                Assert.IsNotNull(sheet);

                var cell = sheet.GetCell(5, 3);
                Assert.IsNotNull(cell);
                Assert.AreEqual(value.ToString(format), cell.GetDateValue().ToString(format));

                DateTime res;
                Assert.AreEqual(true, cell.TryGetValue(out res));
                Assert.AreEqual(value.ToString(format), res.ToString(format));
            }
        }

        [TestMethod]
        public void CellHasValue()
        {
            var filePath = GetFilepath("doc23.xlsx");
            var value = "Aika";
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet();
                var cell = sheet.AddCell(5, 3);
                cell.SetValue(value);

                Assert.AreEqual(value, cell.GetStringValue());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.Worksheets.FirstOrDefault();
                Assert.IsNotNull(sheet);

                var cell = sheet.GetCell(5, 3);
                Assert.IsNotNull(cell);
                Assert.AreEqual(true, cell.HasValue());
                Assert.AreEqual(false, cell.HasFormula());
            }
        }

        [TestMethod]
        public void CellHasFormula()
        {
            var filePath = GetFilepath("doc24.xlsx");
            var formula = "SUM(C1:C4)";
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet();
                var row = sheet.AddRow(3);
                row.AddCellWithValue(15);
                row.AddCellWithValue(4);
                row.AddCellWithValue(-2);
                row.AddCellWithValue(9);

                var cell = row.AddCellWithFormula(5, formula);
                Assert.AreEqual(formula, cell.GetFormula());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.Worksheets.FirstOrDefault();
                Assert.IsNotNull(sheet);

                var cell = sheet.GetCell(5, 3);
                Assert.IsNotNull(cell);
                Assert.AreEqual(false, cell.HasValue());
                Assert.AreEqual(true, cell.HasFormula());
            }
        }

        [TestMethod]
        public void CellHasFormulaGetValue()
        {
            var filePath = GetFilepath("doc25.xlsx");
            var formula = "SUM(C1:C4)";
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet();
                var row = sheet.AddRow(3);
                row.AddCellWithValue(15);
                row.AddCellWithValue(4);
                row.AddCellWithValue(-2);
                row.AddCellWithValue(9);

                var cell = row.AddCellWithFormula(5, formula);
                Assert.AreEqual(formula, cell.GetFormula());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.Worksheets.FirstOrDefault();
                Assert.IsNotNull(sheet);

                var cell = sheet.GetCell(5, 3);
                Assert.IsNotNull(cell);
                Assert.AreEqual(false, cell.HasValue());
                Assert.AreEqual(true, cell.HasFormula());

                try
                {
                    cell.GetIntValue();
                }
                catch (Exception e)
                {
                    Assert.IsInstanceOfType(e, typeof(InvalidOperationException));
                }
            }
        }

        // [TestMethod] // TODO fix
        public void GetDateTimeFromString()
        {
            var filePath = GetFilepath("doc26.xlsx");
            var format = "dd.MM.yyyy hh:mm:ss";
            var value = DateTime.Now.ToString(format);
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet();
                var cell = sheet.AddCellWithValue(5, 3, value);

                Assert.AreEqual(value, cell.GetStringValue());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var time = DateTime.Parse(value);
                var sheet = w.Worksheets.FirstOrDefault();
                Assert.IsNotNull(sheet);

                var cell = sheet.GetCell(5, 3);
                Assert.IsNotNull(cell);
                Assert.AreEqual(value, cell.GetDateValue(format).ToString(format));
                //Assert.AreEqual(time, cell.GetDateValue(format));

                DateTime res;
                Assert.AreEqual(true, cell.TryGetValue(out res, format));
                Assert.AreEqual(value, res.ToString(format));
                //Assert.AreEqual(time, res);
            }
        }
    }
}