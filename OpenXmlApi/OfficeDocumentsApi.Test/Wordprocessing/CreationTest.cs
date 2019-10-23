using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDocumentsApi.Test.TestBases;

namespace OfficeDocumentsApi.Test.Wordprocessing
{
    [TestClass]
    public class CreationTest : WordprocessingTestBase
    {
        public static readonly Random Rnd = new Random();

        [TestMethod]
        public void BasicFile()
        {
            var filename = "doc1.xlsx";
            CleanFilepath(filename);

            var filePath = GetFilepath(filename);
            using (var w = CreateTestee(filePath))
            {
                ;
            }

            Assert.IsTrue(File.Exists(filename));
        }
    }
}
