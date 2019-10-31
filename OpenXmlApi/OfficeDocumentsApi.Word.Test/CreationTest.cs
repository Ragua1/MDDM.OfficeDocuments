using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OfficeDocumentsApi.Word.Test
{
    [TestClass]
    public class CreationTest : WordprocessingTestBase
    {
        public static readonly Random Rnd = new Random();

        [TestMethod]
        public void BasicFile()
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
    }
}
