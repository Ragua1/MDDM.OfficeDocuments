using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDocumentsApi.Word.Enums;

namespace OfficeDocumentsApi.Word.Test
{
    [TestClass]
    public class CreationTest : TestBase
    {
        public static readonly Random Rnd = new Random();

        [TestMethod]
        public void CreateEmptyFile()
        {
            var filename = "doc1.docx";
            CleanFilepath(filename);

            var filePath = filename;// GetFilepath(filename);
            using (var w = CreateTestee(filePath))
            {
                ;
            }

            Assert.IsTrue(File.Exists(filename));
        }

        [TestMethod]
        public void CreateFile()
        {
            var filename = "doc2.docx";
            CleanFilepath(filename);

            var filePath = filename;// GetFilepath(filename);
            using (var w = CreateTestee(filePath))
            {
                var body = w.AddBody();
                body.AddParagraph()
                    .AddText($"Create text on first page - {DateTime.Now:s}")
                    .AddBreak(BreakType.Page);
                
                body = w.AddBody();
                body.AddParagraph()
                    .AddText($"Create text on first page - {DateTime.Now:s}")
                    .AddBreak(BreakType.Page)
                    .AddText($"Create text on second page - {DateTime.Now:s}");
            }

            Assert.IsTrue(File.Exists(filename));
        }
    }
}
