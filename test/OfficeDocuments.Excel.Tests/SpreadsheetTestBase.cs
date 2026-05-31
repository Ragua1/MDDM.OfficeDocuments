using OfficeDocuments.Excel.Interfaces;

namespace OfficeDocuments.Excel.Tests;

public abstract class SpreadsheetTestBase : IDisposable
{
    protected ISpreadsheet CreateNewSpreadsheet(Stream stream) => Spreadsheet.CreateDocument(stream);
    protected ISpreadsheet CreateNewSpreadsheet(string filepath) => new Spreadsheet(filepath, true);

    protected ISpreadsheet OpenExistingSpreadsheet(string filepath) => new Spreadsheet(filepath, false);
    protected ISpreadsheet OpenExistingSpreadsheet(Stream stream) => Spreadsheet.OpenDocument(stream, true);

    protected string GetFilepath(string filename) => TestSettings.GetFilepath(this, filename);

    public void Dispose()
    {
        TestSettings.Cleanup(this);
    }
}