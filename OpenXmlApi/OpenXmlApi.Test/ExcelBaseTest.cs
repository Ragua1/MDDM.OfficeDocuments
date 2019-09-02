using System.IO;

namespace OpenXmlApi.Test
{
    public class ExcelBaseTest
    {
        protected Spreadsheet CreateTestee(Stream stream) => new Spreadsheet(stream, true);
        protected Spreadsheet CreateTestee(string filepath) => new Spreadsheet(filepath, true);

        protected Spreadsheet CreateOpenTestee(string filepath) => new Spreadsheet(filepath, false);
        protected Spreadsheet CreateOpenTestee(Stream stream) => new Spreadsheet(stream, false);

    }
}