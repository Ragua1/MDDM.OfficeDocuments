using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

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

            var wp = new Wordprocessing(path, false);

            var body = wp.GetBody();

            var texts = body.Paragraphs.Select(x => x.GetTexts()).Where(z => !string.IsNullOrEmpty(z)).ToArray();
            
            Assert.IsTrue(texts.Any());
        }
    }
}
