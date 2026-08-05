using OfficeDocuments.Excel.TestKit;
using OfficeDocuments.Excel.TestKit.Validation;

namespace OfficeDocuments.Excel.VerificationTests;

/// <summary>
/// Whole documents built entirely out of the awkward cases (EXCEL-011 phase 6, blind spots B-2
/// through B-6).
/// <para>
/// The integration tier already checks each rule at its own entry point. What only a complete
/// document can answer is whether the escaped forms survive together in one package — a sheet name
/// carrying an ampersand is written into <c>workbook.xml</c>, the same characters in a cell reach
/// the sheet part, and in a comment they reach a third part with its own relationship. Those are
/// three different writers, and a file that Excel opens has to satisfy all of them at once.
/// </para>
/// </summary>
public class EdgeCaseWorkbookTests : SpreadsheetTestBase
{
    /// <summary>
    /// Every place the library writes caller-supplied text, all in one workbook, all containing
    /// characters XML gives special meaning.
    /// </summary>
    [Fact]
    public void MarkupCharactersEverywhere_ProduceAValidWorkbookThatReadsBackUnchanged()
    {
        const string sheetName = "R&D <2026>";
        const string cellValue = "<tag attr=\"v\"> & 'quoted' </tag>";
        const string commentText = "a < b && c > d";
        const string author = "Q&A";

        var filePath = GetFilepath("markup-characters.xlsx");

        using (var spreadsheet = CreateNewSpreadsheet(filePath))
        {
            var worksheet = spreadsheet.AddWorksheet(sheetName);
            worksheet.AddCell(1, 1, cellValue);
            worksheet.AddCell(1, 2, "plain");
            worksheet.GetCell(1, 1)!.SetComment(commentText, author);
            spreadsheet.AddNamedRange("Rand_D", worksheet.GetRange("A1:A2"));
        }

        OpenXmlValidation.AssertValid(filePath);

        using var reopened = OpenExistingSpreadsheet(filePath);
        var reopenedSheet = reopened.GetWorksheet(sheetName);

        Assert.NotNull(reopenedSheet);
        Assert.Equal(cellValue, reopenedSheet.GetCell(1, 1)!.GetStringValue());
        Assert.Equal(commentText, reopenedSheet.GetCell(1, 1)!.GetComment());
    }

    /// <summary>
    /// The dates around Excel's phantom 29 February 1900, written into a real file. The serials
    /// themselves are asserted in the lower tiers; here the question is only whether a workbook
    /// full of them is still a valid document that reopens.
    /// </summary>
    [Fact]
    public void DatesAroundThePhantomLeapDay_ProduceAValidWorkbook()
    {
        var dates = new[]
        {
            new DateTime(1900, 1, 1),
            new DateTime(1900, 2, 28),
            new DateTime(1900, 3, 1),
            new DateTime(2026, 7, 28),
            new DateTime(9999, 12, 31)
        };

        var filePath = GetFilepath("date-boundaries.xlsx");

        using (var spreadsheet = CreateNewSpreadsheet(filePath))
        {
            var worksheet = spreadsheet.AddWorksheet("Dates");
            for (var i = 0; i < dates.Length; i++)
            {
                worksheet.AddCell(1, (uint)i + 1, dates[i]);
            }
        }

        OpenXmlValidation.AssertValid(filePath);

        using var reopened = OpenExistingSpreadsheet(filePath);
        var reopenedSheet = reopened.GetWorksheet("Dates")!;

        for (var i = 0; i < dates.Length; i++)
        {
            Assert.True(reopenedSheet.GetCell(1, (uint)i + 1)!.TryGetValue(out DateTime readBack));
            Assert.Equal(dates[i], readBack);
        }
    }

    /// <summary>
    /// A sheet name at exactly the 31-character limit, which is the value the schema constrains
    /// the attribute to and therefore the one worth proving against a real validator rather than
    /// against the library's own rule table.
    /// </summary>
    [Fact]
    public void SheetNameAtTheLengthLimit_ProducesAValidWorkbook()
    {
        var name = new string('a', 31);
        var filePath = GetFilepath("sheet-name-limit.xlsx");

        using (var spreadsheet = CreateNewSpreadsheet(filePath))
        {
            spreadsheet.AddWorksheet(name).AddCell(1, 1, "value");
        }

        OpenXmlValidation.AssertValid(filePath);

        using var reopened = OpenExistingSpreadsheet(filePath);
        Assert.Contains(name, reopened.GetWorksheetsName());
    }
}
