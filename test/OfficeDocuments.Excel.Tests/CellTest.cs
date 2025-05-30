﻿using System;
using System.Globalization;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDocuments.Excel.Interfaces;
using Color = System.Drawing.Color;
using Styles_Font = OfficeDocuments.Excel.Styles.Font;
using Styles_NumberingFormat = OfficeDocuments.Excel.Styles.NumberingFormat;

namespace OfficeDocuments.Excel.Tests
{
    [TestClass]
    public class CellTest : SpreadsheetTestBase
    {
        [TestMethod]
        public void CreateCell()
        {
            var filePath = GetFilepath("doc1.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var cell = sheet.AddCell();
            Assert.IsNotNull(cell, "New cell cannot be null");
            Assert.IsNotNull(cell.Element, "Cells element cannot be null");
            Assert.IsInstanceOfType(cell, typeof(ICell), "Expected ICell type.");

            Assert.IsTrue(sheet.CurrentRow.Cells.Contains(cell));
            Assert.IsTrue(sheet.CurrentRow.RowIndex == cell.RowIndex);
        }

        [TestMethod]
        public void CreateCellWithValue()
        {
            var filePath = GetFilepath("doc2.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                const string value = "Aloha";
                var cell = sheet.AddCell(value);
                Assert.AreEqual(cell.Value, value);
            }
            Assert.IsTrue(File.Exists(filePath));
        }

        [TestMethod]
        public void SetString()
        {
            var filePath = GetFilepath("doc3.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var value = "Aloha";
            var cell1 = sheet.AddCell(value);
            var cell2 = sheet.AddCell();
            cell2.Value = value;
            Assert.AreEqual(value, cell1.Value, $"Cell value of 'cell.SetValue()' is '{cell1.Value}', expected {value}");
            Assert.AreEqual(value, cell2.Value, $"Cell value of 'cell.SetValue()' is '{cell2.Value}', expected {value}");
            Assert.AreEqual((object)cell1.Value, cell2.Value, $"Cell values are not same");

            Assert.AreEqual(cell1.Style.NumberFormatId, 49, $"Number format is '{cell1.Style.NumberFormatId}', expected 49-'@'");
        }

        [TestMethod]
        public void SetInteger()
        {
            var filePath = GetFilepath("doc4.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var value = 165752313;
            var cell1 = sheet.AddCell(value);
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
            Assert.AreEqual((object)cell1.Value, cell2.Value, $"Cell values are not same");

            Assert.AreEqual(cell1.Style.NumberFormatId, 1, $"Number format is '{cell1.Style.NumberFormatId}', expected 1-'0'");
        }

        [TestMethod]
        public void SetIntegerWithStyle()
        {
            var filePath = GetFilepath("doc5.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var s = w.CreateStyle(numberFormat: new Styles_NumberingFormat("#,##0x"));
            var value = 98435123;
            var cell1 = sheet.AddCell(value, s);
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
            Assert.AreEqual((object)cell1.Value, cell2.Value, $"Cell values are not same");

            var excelUserFormatsIndex = 170;
            Assert.IsTrue(cell1.Style.NumberFormatId >= excelUserFormatsIndex, $"Number format is '{cell1.Style.NumberFormatId}', expected {excelUserFormatsIndex}+");
        }

        [TestMethod]
        public void SetDouble()
        {
            var filePath = GetFilepath("doc6.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var value = 165752313.216546;
            var cell1 = sheet.AddCell(value);
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
            Assert.AreEqual((object)cell1.Value, cell2.Value, $"Cell values are not same");
        }

        [TestMethod]
        public void SetDoubleWithStyle()
        {
            var filePath = GetFilepath("doc7.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var s = w.CreateStyle(numberFormat: new Styles_NumberingFormat("#,##0.##x"));
            var value = 645.541;
            var cell1 = sheet.AddCell(value, s);
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
            Assert.AreEqual((object)cell1.Value, cell2.Value, $"Cell values are not same");

            var excelUserFormatsIndex = 170;
            Assert.IsTrue(cell1.Style.NumberFormatId > excelUserFormatsIndex, $"Number format is '{cell1.Style.NumberFormatId}', expected {excelUserFormatsIndex}+");
        }

        [TestMethod]
        public void SetLong()
        {
            var filePath = GetFilepath("doc100.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var value = 165752313216546;
            var cell1 = sheet.AddCell(value);
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
            Assert.AreEqual((object)cell1.Value, cell2.Value, $"Cell values are not same");
        }

        [TestMethod]
        public void SetLongWithStyle()
        {
            var filePath = GetFilepath("doc101.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var s = w.CreateStyle(numberFormat: new Styles_NumberingFormat("#,##0x"));
            var value = 165752313216546;
            var cell1 = sheet.AddCell(value, s);
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
            Assert.AreEqual((object)cell1.Value, cell2.Value, $"Cell values are not same");

            var excelUserFormatsIndex = 170;
            Assert.IsTrue(cell1.Style.NumberFormatId >= excelUserFormatsIndex, $"Number format is '{cell1.Style.NumberFormatId}', expected {excelUserFormatsIndex}+");
        }

        [TestMethod]
        public void SetDecimal()
        {
            var filePath = GetFilepath("doc102.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var value = 165752313216546.6516511m;
            var cell1 = sheet.AddCell(value);
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
            Assert.AreEqual((object)cell1.Value, cell2.Value, "Cell values are not same");
        }

        [TestMethod]
        public void SetDecimalWithStyle()
        {
            var filePath = GetFilepath("doc103.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var s = w.CreateStyle(numberFormat: new Styles_NumberingFormat("#,##0.##0x"));
            var value = 16575231321654.6565465426m;
            var cell1 = sheet.AddCell(value, s);
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
            Assert.AreEqual((object)cell1.Value, cell2.Value, $"Cell values are not same");

            var excelUserFormatsIndex = 170;
            Assert.IsTrue(
                cell1.Style.NumberFormatId >= excelUserFormatsIndex, $"Number format is '{cell1.Style.NumberFormatId}', expected {excelUserFormatsIndex}+"
            );
        }

        [TestMethod]
        public void SetDate()
        {
            var filePath = GetFilepath("doc8.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var value = DateTime.Now;
            var cell1 = sheet.AddCell(value);
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
            Assert.AreEqual((object)cell1.Value, cell2.Value, $"Cell values are not same");

            Assert.AreEqual(cell1.Style.NumberFormatId, 14, $"Number format is '{cell1.Style.NumberFormatId}', expected 14-'d/m/yyyy'");
        }

        [TestMethod]
        public void SetDateWithStyle()
        {
            var filePath = GetFilepath("doc9.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var s = w.CreateStyle(numberFormat: new Styles_NumberingFormat("d/m/yyyy H:mm:ss"));
            var value = DateTime.Now;
            var cell1 = sheet.AddCell(value, s);
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
            Assert.AreEqual((object)cell1.Value, cell2.Value, $"Cell values are not same");

            var excelUserFormatsIndex = 170;
            Assert.IsTrue(
                cell1.Style.NumberFormatId >= excelUserFormatsIndex,
                $"Number format is '{cell1.Style.NumberFormatId}', expected '{excelUserFormatsIndex}'+"
            );
        }

        [TestMethod]
        public void SetBoolean()
        {
            var filePath = GetFilepath("doc10.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var value = true;
            var cell1 = sheet.AddCell(value);
            var cell2 = sheet.AddCell();
            cell2.SetValue(value);
            Assert.AreEqual(cell1.Value, value.ToString(), $"Cell value of 'cell.SetValue()' is '{cell1.Value}', expected {value}");
            Assert.AreEqual(cell2.Value, value.ToString(), $"Cell value of 'cell.SetValue()' is '{cell2.Value}', expected {value}");
            Assert.AreEqual((object)cell1.Value, cell2.Value, $"Cell values are not same");
        }

        [TestMethod]
        public void SetFormula()
        {
            var filePath = GetFilepath("doc11.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var formula = "SUM(A1:A5)";
            var cell = sheet.AddCell();
            cell.SetFormula(formula);
            Assert.AreEqual(cell.Element.CellFormula.Text, formula, $"Cell formula is '{cell.Element.CellFormula.Text}', expected {formula}");
        }

        [TestMethod]
        public void SumInRangeFormula()
        {
            int[] numbers = new int[] { 1, 2, 3 };
            int sum = 0;
            foreach (int num in numbers)
            {
                sum += num;
            }

            var filePath = GetFilepath("doc30.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var formula = "SUM(B1:D1)";

            var cell1 = sheet.AddCellWithFormula(formula);
            var cell2 = sheet.AddCell(numbers[0]);
            var cell3 = sheet.AddCell(numbers[1]);
            var cell4 = sheet.AddCell(numbers[2]);

            int value = cell1.GetFormulaValue();

            Assert.AreEqual(value, sum);
        }

        [TestMethod]
        public void SumInRangeFormula2()
        {
            int[] numbers = new int[] { 1, 2, 3 };
            int sum = 0;
            foreach (int num in numbers)
            {
                sum += num;
            }

            var filePath = GetFilepath("doc31.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var formula = "SUM(B1:H1)";

            var cell1 = sheet.AddCellWithFormula(formula);
            var cell2 = sheet.AddCell(numbers[0]);
            var cell3 = sheet.AddCell(numbers[1]);
            var cell4 = sheet.AddCell(numbers[2]);

            int value = cell1.GetFormulaValue();

            Assert.AreEqual(value, sum);
        }

        [TestMethod]
        public void SumInRangeFormula3()
        {
            int[] numbers = new int[] { 1, 2, 3 };
            int sum = 0;
            foreach (int num in numbers)
            {
                sum += num;
            }

            string text = "Lorem Ipsum";

            var filePath = GetFilepath("doc32.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var formula = "SUM(B1:D1)";

            var cell1 = sheet.AddCellWithFormula(formula);
            var cell2 = sheet.AddCell(numbers[0]);
            var cell3 = sheet.AddCell(numbers[1]);
            var cell4 = sheet.AddCell(text);

            Assert.ThrowsException<ArgumentException>(() => cell1.GetFormulaValue());
            //int value = cell1.GetFormulaValue();

            //Assert.AreEqual(value, sum);
        }

        [TestMethod]
        public void SumInRangeFormula4()
        {
            int[] numbers = new int[] { 1, 2, 3, 4 };
            int sum = 0;
            foreach (int num in numbers)
            {
                sum += num;
            }

            string formula2 = "SUM(E1:F1)";

            var filePath = GetFilepath("doc33.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var formula = "SUM(B1:D1)";

            var cell1 = sheet.AddCellWithFormula(formula);
            var cell2 = sheet.AddCell(numbers[0]);
            var cell3 = sheet.AddCell(numbers[1]);
            var cell4 = sheet.AddCell(formula2);
            var cell5 = sheet.AddCell(numbers[2]);
            var cell6 = sheet.AddCell(numbers[3]);


            Assert.ThrowsException<ArgumentException>(() => cell1.GetFormulaValue());
            //int value = cell1.GetFormulaValue();

            //Assert.AreEqual(value, sum);
        }

        [TestMethod]
        public void SumInRangeFormula5()
        {
            int[] numbers = new int[] { 1, 2, 3, 4 };
            int sum = 0;
            foreach (int num in numbers)
            {
                sum += num;
            }

            string formula2 = "SUM(E1:F1)";

            var filePath = GetFilepath("doc34.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var formula = "SUM(B1:D1)";

            var cell1 = sheet.AddCellWithFormula(formula);
            var cell2 = sheet.AddCell(numbers[0]);
            var cell3 = sheet.AddCell(numbers[1]);
            var cell4 = sheet.AddCellWithFormula(formula2);
            var cell5 = sheet.AddCell(numbers[2]);
            var cell6 = sheet.AddCell(numbers[3]);


            int value = cell1.GetFormulaValue();

            Assert.AreEqual(value, sum);
        }

        [TestMethod]
        public void CellInheritStyleFromRow()
        {
            var filePath = GetFilepath("doc12.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var s = w.CreateStyle(new Styles_Font { Color = Color.DarkGoldenrod });
            var row = sheet.AddRow(s);
            var cell = row.AddCell();

            Assert.AreEqual(
                row.Style.StyleIndex, cell.Style.StyleIndex,
                $"Cell not inherit style. Row style '{row.Style.StyleIndex}', cell stylw {cell.Style.StyleIndex}"
            );
        }

        [TestMethod]
        public void CellInheritStyleFromSheet()
        {
            var filePath = GetFilepath("doc13.xlsx");
            using var w = CreateTestee(filePath);
            var s = w.CreateStyle(new Styles_Font { Color = Color.DarkGoldenrod });
            var sheet = w.AddWorksheet("Sheet 1", s);
            var cell = sheet.AddCell();

            Assert.AreEqual(
                (object)sheet.Style.StyleIndex, cell.Style.StyleIndex,
                $"Cell not inherit style. Sheet style '{sheet.Style.StyleIndex}', cell style {cell.Style.StyleIndex}"
            );
        }

        [TestMethod]
        public void CellHasCorrectIndexes1()
        {
            var filePath = GetFilepath("doc14.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");
            var cell = sheet.AddCellOnIndex(5, 3);

            Assert.AreEqual((uint)5, cell.ColumnIndex);
            Assert.AreEqual((uint)3, cell.RowIndex);
            Assert.AreEqual("E3", cell.CellReference);
        }

        [TestMethod]
        public void CellHasCorrectIndexes2()
        {
            var filePath = GetFilepath("doc15.xlsx");
            using var w = CreateTestee(filePath);
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

        [TestMethod]
        public void CellHasCorrectIndexes3()
        {
            var filePath = GetFilepath("doc16.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var cell = sheet.AddCellOnIndex(5, 3);

                Assert.AreEqual((uint)5, cell.ColumnIndex);
                Assert.AreEqual((uint)3, cell.RowIndex);
                Assert.AreEqual("E3", cell.CellReference);
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
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
                var cell = sheet.AddCellOnIndex(5, 3);
                cell.SetValue(true);

                Assert.AreEqual(true, cell.GetBoolValue());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
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
                var cell = sheet.AddCellOnIndex(5, 3);
                cell.SetValue(int.MaxValue);

                Assert.AreEqual(int.MaxValue, cell.GetIntValue());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
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
                var cell = sheet.AddCellOnIndex(5, 3);
                cell.SetValue(long.MaxValue);

                Assert.AreEqual(long.MaxValue, cell.GetLongValue());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
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
                var cell = sheet.AddCellOnIndex(5, 3);
                cell.SetValue(value);

                Assert.AreEqual(value, cell.GetDoubleValue());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
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
                var cell = sheet.AddCellOnIndex(5, 3);
                cell.SetValue(decimal.MaxValue);

                Assert.AreEqual(decimal.MaxValue, cell.GetDecimalValue());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
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
                var cell = sheet.AddCellOnIndex(5, 3);
                cell.SetValue(value);

                Assert.AreEqual(value, cell.GetStringValue());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
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
                var cell = sheet.AddCellOnIndex(5, 3);
                cell.SetValue(value);

                Assert.AreEqual(value.ToString(format), cell.GetDateValue().ToString(format));
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
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
                var cell = sheet.AddCellOnIndex(5, 3);
                cell.SetValue(value);

                Assert.AreEqual(value, cell.GetStringValue());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
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
                row.AddCell(15);
                row.AddCell(4);
                row.AddCell(-2);
                row.AddCell(9);

                var cell = row.AddCellWithFormula(5, formula);
                Assert.AreEqual(formula, cell.GetFormula());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
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
                row.AddCell(15);
                row.AddCell(4);
                row.AddCell(-2);
                row.AddCell(9);

                var cell = row.AddCellWithFormula(5, formula);
                Assert.AreEqual(formula, cell.GetFormula());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
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
                var cell = sheet.AddCell(5, 3, value);

                Assert.AreEqual(value, cell.GetStringValue());
            }

            using (var w = CreateOpenTestee(filePath))
            {
                var time = DateTime.Parse(value);
                var sheet = w.GetWorksheet(w.GetWorksheetsName().First());
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

        [TestMethod]
        public void CreateMultipleCellsWithValue()
        {
            var filePath = GetFilepath("doc27.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                for (int i = 0; i < 10; i++)
                {
                    var cell = sheet.AddCell(i);
                    Assert.AreEqual(cell.GetIntValue(), i);
                }
            }
            Assert.IsTrue(File.Exists(filePath));
        }

        [TestMethod]
        public void CreateMultipleCellsWithStyle()
        {
            var filePath = GetFilepath("doc28.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                for (int i = 0; i < 10; i++)
                {
                    var s = w.CreateStyle(numberFormat: new Styles_NumberingFormat("#,##0x"));
                    var value = i;
                    var cell1 = sheet.AddCell(value, s);
                    var cell2 = sheet.AddCell(s);
                    cell2.SetValue(value);

                    Assert.AreEqual(cell1.GetIntValue(), value);
                    Assert.AreEqual(cell2.GetIntValue(), value);
                    Assert.AreEqual(cell1.GetIntValue(), cell2.GetIntValue());
                }
            }
            Assert.IsTrue(File.Exists(filePath));
        }

        [TestMethod]
        public void CreateRandomCellsInRange()
        {
            var filePath = GetFilepath("doc29.xlsx");
            using var w = CreateTestee(filePath);
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
                        Assert.AreEqual(cell.GetIntValue(), j);
                    }
                    else
                    {
                        var cell = sheet.AddCell();
                        Assert.IsNotNull(cell);
                    }
                }
            }
        }

        [TestMethod]
        public void CountCellsWithValueInRange()
        {
            var filePath = GetFilepath("doc35.xlsx");
            using var w = CreateTestee(filePath);
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

            Assert.AreEqual(3, sum, $"Number of cells with value is not correct. Expected 3, actual {sum}");
        }

        [TestMethod]

        // Use string in formula as if argument
        public void CountCellsByArgument()
        {
            var filePath = GetFilepath("doc36.xlsx");
            using var w = CreateTestee(filePath);
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

            Assert.AreEqual(3, sum, $"Number of cells with argument value is not correct. Expected 3, actual {sum}");
        }

        [TestMethod]

        // Use value of argumented cell as if argument
        public void CountCellsByArgument2()
        {
            var filePath = GetFilepath("doc37.xlsx");
            using var w = CreateTestee(filePath);
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

            Assert.AreEqual(3, sum, $"Number of cells with argument value is not correct. Expected 3, actual {sum}");
        }

        [TestMethod]

        // Use value of empty cell as if argument
        public void CountCellsByArgument3()
        {
            var filePath = GetFilepath("doc38.xlsx");
            using var w = CreateTestee(filePath);
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

            Assert.AreEqual(3, sum, $"Number of cells with argument value is not correct. Expected 3, actual {sum}");
        }

        [TestMethod]

        // Calculate median of an odd row of numbers
        public void GetMedian()
        {
            var filePath = GetFilepath("doc39.xlsx");
            using var w = CreateTestee(filePath);
            var sheet = w.AddWorksheet("Sheet 1");

            var formula = "MEDIAN(B1:F1)";

            var cell1 = sheet.AddCellWithFormula(formula);

            var cell2 = sheet.AddCell(1);
            var cell3 = sheet.AddCell(5);
            var cell4 = sheet.AddCell(7);
            var cell5 = sheet.AddCell(9);
            var cell6 = sheet.AddCell(2);

            var median = cell1.GetFormulaValue();
            var expcted = 5;

            Assert.AreEqual(expcted, median, $"Calculated median is not correct. Expected {expcted}, actual {median}");
        }

        [TestMethod]

        // Calculate median of an even row of numbers
        public void GetMedian2()
        {
            var filePath = GetFilepath("doc40.xlsx");
            using var w = CreateTestee(filePath);
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
            var expcted = 6;

            Assert.AreEqual(expcted, median, $"Calculated median is not correct. Expected {expcted}, actual {median}");
        }

        [TestMethod]

        // Calculate median of a row of numbers with strings
        public void GetMedian3()
        {
            var filePath = GetFilepath("doc41.xlsx");
            using var w = CreateTestee(filePath);
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
            var expcted = 3;

            Assert.AreEqual(expcted, median, $"Calculated median is not correct. Expected {expcted}, actual {median}");
        }
    }
}