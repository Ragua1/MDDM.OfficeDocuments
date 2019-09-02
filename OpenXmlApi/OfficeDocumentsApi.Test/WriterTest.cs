using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDocumentsApi.Emums;
using OfficeDocumentsApi.Styles;
using Color = System.Drawing.Color;

namespace OfficeDocumentsApi.Test
{
    [TestClass]
    public class WriterTest : ExcelBaseTest
    {
       


        [TestMethod]
        public void CreateNewFile()
        {
            var filePath = GetFilepath("doc1.xlsx");
            DeleteFile(filePath);

            using (var writer = CreateTestee(filePath))
            {
                ;
            }

            Assert.IsTrue(File.Exists(filePath), $"File not exist on filePath: '{filePath}'");
        }

        [TestMethod]
        public void CreateStyleWithDefaultSettings()
        {
            var filePath = GetFilepath("doc2.xlsx");
            DeleteFile(filePath);

            using (var writer = CreateTestee(filePath))
            {
                var style = writer.CreateStyle();

                Assert.AreEqual(0, style.FontId);
                Assert.AreEqual(0, style.FillId);
                Assert.AreEqual(0, style.BorderId);
                Assert.AreEqual(0, style.NumberFormatId);

                Assert.AreEqual((uint)0, style.StyleIndex);
            }
        }

        [TestMethod]
        public void CreateStyle()
        {
            var filePath = GetFilepath("doc3.xlsx");
            DeleteFile(filePath);

            using (var writer = CreateTestee(filePath))
            {
                var style = writer.CreateStyle(
                    new Font { FontSize = 12, Color = Color.Aqua, FontName = FontNameValues.Tahoma },
                    new Fill(Color.Brown),
                    new Border { Left = BorderStyleValues.Double, Right = BorderStyleValues.Double }
                );

                Assert.IsTrue(style.FontId > 0);
                Assert.IsTrue(style.FillId > 0);
                Assert.IsTrue(style.BorderId > 0);
                Assert.IsTrue(style.NumberFormatId == 0);

                Assert.IsTrue(style.StyleIndex > 0);
            }
        }

        [TestMethod]
        public void NotCreateFileOnNonExistDirectory()
        {
            var filePath = GetFilepath("doc4.xlsx");
            DeleteDirectory(Path.GetDirectoryName(filePath));

            try
            {
                using (var writer = CreateTestee(filePath))
                {
                    Assert.Fail();
                }
            }
            catch (Exception ex)
            {
                Assert.IsInstanceOfType(ex, typeof(DirectoryNotFoundException));
            }

            Assert.IsFalse(File.Exists(filePath), $"File should not exist on filePath: '{filePath}'");
        }

        [TestMethod]
        public void GetExistedWorksheet()
        {
            var filePath = GetFilepath("doc5.xlsx");
            DeleteFile(filePath);

            using (var writer = CreateTestee(filePath))
            {
                var sheetName = "Test1";
                var sheet1 = writer.AddWorksheet(sheetName);

                var sheet2 = writer.GetWorksheet(sheetName);

                Assert.IsNotNull(sheet2, $"Spreadsheet cannot find sheet '{sheetName}'");
                Assert.AreSame(sheet1, sheet2, "Sheets are not same!");
            }
        }

        [TestMethod]
        public void OpenExistedDocument()
        {
            var filePath = GetFilepath("doc6.xlsx");
            DeleteFile(filePath);

            using (var writer = CreateTestee(filePath))
            {
                writer.AddWorksheet("Test1");
            }

            using (var writer = CreateOpenTestee(filePath))
            {
                Assert.IsTrue(writer.Worksheets.Any());
            }
        }

        [TestMethod]
        public void OpenTryFindNonExistedWorksheet()
        {
            var filePath = GetFilepath("doc7.xlsx");
            DeleteFile(filePath);

            using (var writer = CreateTestee(filePath))
            {
                writer.AddWorksheet("Test1");
            }

            using (var writer = CreateOpenTestee(filePath))
            {
                Assert.IsTrue(writer.Worksheets.Any());
                Assert.IsNull(writer.GetWorksheet("Test"));
            }
        }

        [TestMethod]
        public void CreateDocumentToStream()
        {
            var stream = new MemoryStream();

            using (var writer = CreateTestee(stream))
            {
                writer.AddWorksheet("Test1");
            }
        }

        [TestMethod]
        public void OpenDocumentFromStream()
        {
            var stream = new MemoryStream();

            using (var writer = CreateTestee(stream))
            {
                writer.AddWorksheet("Test1");
            }

            using (var writer = CreateOpenTestee(stream))
            {
                Assert.IsTrue(writer.Worksheets.Any());
            }
        }

        private string GetFilepath(string filename)
        {
            return TestSettings.GetFilepath(this, filename);
        }

        private void DeleteFile(string filePath)
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }

        private void DeleteDirectory(string dirPath)
        {
            if (Directory.Exists(dirPath))
            {
                Directory.Delete(dirPath, true);
            }
        }
    }
}