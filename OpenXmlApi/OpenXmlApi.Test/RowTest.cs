using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlApi.Interfaces;
using OpenXmlApi.Styles;
using Color = System.Drawing.Color;

namespace OpenXmlApi.Test
{
    [TestClass]
    public class RowTest : ExcelBaseTest
    {
        [TestMethod]
        public void CreateRow()
        {
            var filePath = GetFilepath("doc1.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var row = sheet.AddRow();
                Assert.IsNotNull(row, "New row cannot be null");
                Assert.IsNotNull(row.Element, "Row element cannot be null");
                Assert.IsInstanceOfType(row, typeof(IRow), "Expected ICell type.");

                Assert.IsTrue(sheet.Rows.Contains(row));
                Assert.IsTrue(sheet.CurrentRow.RowIndex == row.RowIndex);
                Assert.IsTrue(row.CurrentCell == null);
            }
        }

        [TestMethod]
        public void CreateRowWithStyle()
        {
            var filePath = GetFilepath("doc2.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var s = w.CreateStyle(new Font { Color = Color.Coral }, new Fill(Color.Black));
                var row = sheet.AddRow(s);

                Assert.IsTrue(row.Style.StyleIndex > 0);
            }
        }

        [TestMethod]
        public void CreateRowOnSpecificRowIndex()
        {
            var filePath = GetFilepath("doc3.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var row = sheet.AddRow(5);

                Assert.AreEqual((uint)5, row.RowIndex);
            }
        }

        [TestMethod]
        public void CreateRowOnWrongRowIndex()
        {
            var filePath = GetFilepath("doc4.xlsx");
            try
            {
                using (var w = CreateTestee(filePath))
                {
                    var sheet = w.AddWorksheet("Sheet 1");
                    sheet.AddRow(0);
                }
            }
            catch (Exception e)
            {
                Assert.IsInstanceOfType(e, typeof(ArgumentException));
            }
        }

        [TestMethod]
        public void CreateRowAndAddCell()
        {
            var filePath = GetFilepath("doc5.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var row = sheet.AddRow();
                var cell1 = row.AddCell();
                var cell2 = row.AddCell(3);

                Assert.AreEqual((uint)1, cell1.ColumnIndex);
                Assert.AreEqual((uint)3, cell2.ColumnIndex);
            }
        }

        [TestMethod]
        public void CreateRowAndAddCellWithValue()
        {
            var filePath = GetFilepath("doc6.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var value = "Alea iacta est";
                var row = sheet.AddRow();
                var cell1 = row.AddCellWithValue(value);
                var cell2 = row.AddCellWithValue(3, value);

                Assert.AreEqual((uint)1, cell1.ColumnIndex);
                Assert.AreEqual(value, cell1.Value);

                Assert.AreEqual((uint)3, cell2.ColumnIndex);
                Assert.AreEqual(value, cell2.Value);

                value = "Sumilian Eri Lopte";
                cell1 = row.AddCellWithValue(1, value);
                Assert.AreEqual((uint)1, cell1.ColumnIndex);
                Assert.AreEqual(value, cell1.Value);
            }
        }

        [TestMethod]
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

                Assert.AreEqual((uint)1, cell.ColumnIndex);
                Assert.AreEqual(value, cell.Element.CellFormula.Text);

                value = "Sum(B1:B2)";
                cell = row.AddCellWithFormula(1, value);
                Assert.AreEqual((uint)1, cell.ColumnIndex);
                Assert.AreEqual(value, cell.Element.CellFormula.Text);
            }
        }

        [TestMethod]
        public void CreateRowWithValueOnWrongRowIndex()
        {
            var filePath = GetFilepath("doc8.xlsx");
            try
            {
                using (var w = CreateTestee(filePath))
                {
                    var sheet = w.AddWorksheet("Sheet 1");
                    sheet.AddCellWithValue(0, 0);
                }
            }
            catch (Exception e)
            {
                Assert.IsInstanceOfType(e, typeof(ArgumentException));
            }
        }

        [TestMethod]
        public void CreateRowWithFormulaOnWrongRowIndex()
        {
            var filePath = GetFilepath("doc9.xlsx");
            try
            {
                using (var w = CreateTestee(filePath))
                {
                    var sheet = w.AddWorksheet("Sheet 1");
                    sheet.AddCellWithFormula(0, "0");
                }
            }
            catch (Exception e)
            {
                Assert.IsInstanceOfType(e, typeof(ArgumentException));
            }
        }

        [TestMethod]
        public void CreateRowAndCellOnRange()
        {
            var filePath = GetFilepath("doc10.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var row = sheet.AddRow();
                row.AddCellOnRange(2, 4);

                Assert.IsTrue(row.GetCell(2) != null);
                Assert.IsTrue(row.GetCell(3) != null);
                Assert.IsTrue(row.GetCell(4) != null);
            }
        }

        [TestMethod]
        public void CreateRowAndCellOnRangeOnWrongColumnIndex()
        {
            var filePath = GetFilepath("doc11.xlsx");
            try
            {
                using (var w = CreateTestee(filePath))
                {
                    var sheet = w.AddWorksheet("Sheet 1");
                    var row = sheet.AddRow();
                    row.AddCellOnRange(0, 4);
                }
            }
            catch (Exception e)
            {
                Assert.IsInstanceOfType(e, typeof(ArgumentException));
            }
        }

        [TestMethod]
        public void CreateRowAndCellOnWrongRange()
        {
            var filePath = GetFilepath("doc12.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var row = sheet.AddRow();
                var cell = row.AddCellOnRange(5, 4);

                Assert.IsNull(cell);
            }
        }

        [TestMethod]
        public void CreateRowAndCellOnBigRange()
        {
            var filePath = GetFilepath("doc13.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var row = sheet.AddRow();
                var cell = row.AddCellOnRange(2, 611);

                Assert.IsNotNull(cell);
            }
        }

        [TestMethod]
        public void GetCellByName()
        {
            var filePath = GetFilepath("doc14.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var value = "Alea iacta est";
                var row = sheet.AddRow();
                row.AddCellWithValue(value);
                row.AddCellWithValue(3, value);

                var cell1 = row.GetCell("A");
                Assert.IsNotNull(cell1);
                Assert.AreEqual((uint)1, cell1.ColumnIndex);
                Assert.AreEqual(value, cell1.Value);

                var cell2 = row.GetCell("C");
                Assert.IsNotNull(cell2);
                Assert.AreEqual((uint)3, cell2.ColumnIndex);
                Assert.AreEqual(value, cell2.Value);

                value = "Sumilian Eri Lopte";
                row.AddCellWithValue(1, value);

                cell1 = row.GetCell("A");
                Assert.AreEqual((uint)1, cell1.ColumnIndex);
                Assert.AreEqual(value, cell1.Value);
            }
        }

        private string GetFilepath(string filename)
        {
            return TestSettings.GetFilepath(this, filename);
        }
    }
}