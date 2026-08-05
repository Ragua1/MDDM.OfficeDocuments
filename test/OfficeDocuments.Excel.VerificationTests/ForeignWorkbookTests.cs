using OfficeDocuments.Excel.TestKit;
using OfficeDocuments.Excel.TestKit.Validation;
using OfficeDocuments.Excel.VerificationTests.Properties;

namespace OfficeDocuments.Excel.VerificationTests;

/// <summary>
/// Reads workbooks this library did not write.
/// </summary>
/// <remarks>
/// Every other test round-trips through our own writer, which only ever proves the reader
/// understands our own dialect. <c>Example_2.xlsx</c> comes from Microsoft Excel and stores its
/// text in <c>sharedStrings.xml</c> (<c>t="s"</c> cells) rather than as inline strings, so it is
/// the only coverage the shared-string read path has.
/// </remarks>
public class ForeignWorkbookTests : SpreadsheetTestBase
{
    /// <summary>
    /// An expandable copy of the fixture. <c>new MemoryStream(byte[])</c> is fixed-size, and the
    /// library opens a stream for editing, so the package fails to flush on dispose.
    /// </summary>
    private static MemoryStream OpenExcelAuthoredWorkbook()
    {
        var stream = new MemoryStream();
        stream.Write(Resources.Example_2, 0, Resources.Example_2.Length);
        stream.Position = 0;

        return stream;
    }

    [Fact]
    public void ExcelAuthoredWorkbook_SharedStrings_AreReadAsText()
    {
        using var spreadsheet = OpenExistingSpreadsheet(OpenExcelAuthoredWorkbook());
        var worksheet = spreadsheet.GetWorksheet("Sheet1");

        Assert.NotNull(worksheet);
        Assert.Equal("Seznam hráčů minigolfu", worksheet.GetCellByReference("B1")?.GetStringValue());
        Assert.Equal("Hráč", worksheet.GetCellByReference("A2")?.GetStringValue());
        Assert.Equal("Jamka 1", worksheet.GetCellByReference("B2")?.GetStringValue());
        Assert.Equal("Skóre", worksheet.GetCellByReference("G2")?.GetStringValue());
    }

    [Fact]
    public void ExcelAuthoredWorkbook_NonAsciiText_SurvivesTheReadPath()
    {
        using var spreadsheet = OpenExistingSpreadsheet(OpenExcelAuthoredWorkbook());
        var worksheet = spreadsheet.GetWorksheet("Sheet1");

        // Czech diacritics round-trip only if the whole chain stays UTF-8.
        Assert.Equal("Hráč", worksheet?.GetCellByReference("A2")?.GetStringValue());
    }

    [Fact]
    public void ExcelAuthoredWorkbook_IsListedByName()
    {
        using var spreadsheet = OpenExistingSpreadsheet(OpenExcelAuthoredWorkbook());

        Assert.Equal(["Sheet1"], spreadsheet.GetWorksheetsName());
    }

    [Fact]
    public void ExcelAuthoredWorkbook_CanBeExtendedAndStaysValid()
    {
        var stream = OpenExcelAuthoredWorkbook();

        using (var spreadsheet = OpenExistingSpreadsheet(stream))
        {
            var worksheet = spreadsheet.GetWorksheet("Sheet1");
            Assert.NotNull(worksheet);

            var row = worksheet.AddRow();
            row.AddCell("Nový hráč");
            row.AddCell(3);

            spreadsheet.Close();
        }

        OpenXmlValidation.AssertValid(stream);
    }
}
