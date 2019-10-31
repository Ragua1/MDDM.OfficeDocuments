using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDocumentsApi.Excel.Enums;
using OfficeDocumentsApi.Excel.Interfaces;
using OfficeDocumentsApi.Excel.Styles;
using Color = System.Drawing.Color;

namespace OfficeDocumentsApi.Excel.Test
{
    [TestClass]
    public class WorksheetTest : SpreadsheetTestBase
    {
        [TestMethod]
        public void CreateCellOnWrongColumnIndex()
        {
            var filePath = GetFilepath("doc1.xlsx");
            try
            {
                using (var w = CreateTestee(filePath))
                {
                    var sheet = w.AddWorksheet("Sheet 1");
                    sheet.AddCell(0);
                }
            }
            catch (Exception ex)
            {
                Assert.IsInstanceOfType(ex, typeof(ArgumentException));
            }
        }

        [TestMethod]
        public void CreateCellOnWrongRowIndex()
        {
            var filePath = GetFilepath("doc2.xlsx");
            try
            {
                using (var w = CreateTestee(filePath))
                {
                    var sheet = w.AddWorksheet("Sheet 1");
                    sheet.AddCell(5, 0);
                }
            }
            catch (Exception ex)
            {
                Assert.IsInstanceOfType(ex, typeof(ArgumentException));
            }
        }

        [TestMethod]
        public void CreateCellOnSpecificColumnIndex()
        {
            var filePath = GetFilepath("doc3.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var cell = sheet.AddCell(5);
                Assert.IsNotNull(cell, "New cell cannot be null");
                Assert.IsNotNull(cell.Element, "Cells element cannot be null");
                Assert.IsInstanceOfType(cell, typeof(ICell), "Expected ICell type.");
            }
        }

        [TestMethod]
        public void CreateCellOnSpecificRowIndexAndColumnIndex()
        {
            var filePath = GetFilepath("doc4.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var cell = sheet.AddCell(5, 4);
                Assert.IsNotNull(cell, "New cell cannot be null");
                Assert.IsNotNull(cell.Element, "Cells element cannot be null");
                Assert.IsInstanceOfType(cell, typeof(ICell), "Expected ICell type.");
            }
        }

        [TestMethod]
        public void CreateCellWithStyle()
        {
            var filePath = GetFilepath("doc5.xlsx");
            using (var w = CreateTestee(filePath))
            {
                var sheet = w.AddWorksheet("Sheet 1");
                var s = w.CreateStyle(new Font { Color = Color.Blue }, new Fill(Color.BurlyWood), new Border(BorderStyleValues.Hair), new NumberingFormat("0"));
                var cell = sheet.AddCell(s);
                Assert.IsTrue(cell.Style.FontId > 0);
                Assert.IsTrue(cell.Style.FillId > 0);
                Assert.IsTrue(cell.Style.BorderId > 0);
                Assert.IsTrue(cell.Style.NumberFormatId > 0);
                Assert.IsTrue(cell.Style.StyleIndex > 0);
            }
        }
    }
}