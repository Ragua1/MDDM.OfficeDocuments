using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDocumentsApi.Excel.Interfaces;

namespace OfficeDocumentsApi.Excel.Test
{
    [TestClass]
    public class ReaderTest
    {
        [TestMethod]
        public void CreateNewFile()
        {
        }

        //[TestMethod]
        public void ReadFile()
        {
            var path = @"C:\Users\Martin\Documents\Sešit1.xlsx";

            using ISpreadsheet ss = new Spreadsheet(path);

            var ws = ss.GetWorksheet(ss.GetWorksheetsName().First());

            ss.AddTable("", null, null, null);
            ;
        }
    }
}
