using OfficeDocuments.Excel.Interfaces;

namespace OfficeDocuments.Excel.Tests;

public class ReaderTest
{
    [Fact]
    public void CreateNewFile()
    {
        // Empty test
    }

    // Commented out test remains commented out
    //[Fact]
    public void ReadFile()
    {
        var path = @"C:\Users\Martin\Documents\Sešit1.xlsx";

        using ISpreadsheet ss = new Spreadsheet(path);

        var ws = ss.GetWorksheet(ss.GetWorksheetsName().First());

        ss.AddTable("", null, null, null);
    }
}