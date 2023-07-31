using System.IO;
using OfficeDocumentsApi.Excel.Interfaces;

namespace OfficeDocumentsApi.Excel.Test
{
    public class SpreadsheetTestBase
    {
        protected ISpreadsheet CreateTestee(Stream stream) => Spreadsheet.CreateDocument(stream); // new Spreadsheet(stream, true);
        protected ISpreadsheet CreateTestee(string filepath) => new Spreadsheet(filepath, true);

        protected ISpreadsheet CreateOpenTestee(string filepath) => new Spreadsheet(filepath, false);
        protected ISpreadsheet CreateOpenTestee(Stream stream) => Spreadsheet.OpenDocument(stream, false);

        protected string GetFilepath(string filename) => TestSettings.GetFilepath(this, filename);
    }
}
