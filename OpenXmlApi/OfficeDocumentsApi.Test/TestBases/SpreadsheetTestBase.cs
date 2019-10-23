using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeDocumentsApi.Test.TestBases
{
    public class SpreadsheetTestBase
    {
        protected Spreadsheet CreateTestee(Stream stream) => new Spreadsheet(stream, true);
        protected Spreadsheet CreateTestee(string filepath) => new Spreadsheet(filepath, true);

        protected Spreadsheet CreateOpenTestee(string filepath) => new Spreadsheet(filepath, false);
        protected Spreadsheet CreateOpenTestee(Stream stream) => new Spreadsheet(stream, false);

        protected string GetFilepath(string filename) => TestSettings.GetFilepath(this, filename);
    }
}
