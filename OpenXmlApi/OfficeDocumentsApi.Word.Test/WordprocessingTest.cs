using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDocumentsApi.Word.Interfaces;

namespace OfficeDocumentsApi.Word.Test
{
    [TestClass]
    public class WordprocessingTest : TestBase
    {
        //[TestMethod]
        public void ReadDocumentTest()
        {
            var path = "Resources/Rozsudek_priloha_6.docx";
            Assert.IsTrue(File.Exists(path));

            using IWordprocessing wp = new Wordprocessing(path, false);

            var body = wp.GetBody();

            var texts = body.Paragraphs.Select(x => x.GetTextElements()).Where(x => x.Any()).ToArray();
            
            Assert.IsTrue(texts.Any());

            wp.Close(false);
        }
    }
}
