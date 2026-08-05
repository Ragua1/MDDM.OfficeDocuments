using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Options;
using OfficeDocuments.Excel.Styles;
using OfficeDocuments.Excel.TestKit;
using OfficeDocuments.Excel.TestKit.Validation;
using Color = System.Drawing.Color;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.VerificationTests;

public class WorkbookRoundtripTests : SpreadsheetTestBase
{
    [Fact]
    public void WorkbookRoundtrip_BasicValuesAndFormula_AreReadableAfterReopen()
    {
        var filePath = GetFilepath("reader-basic-roundtrip.xlsx");

        using (var spreadsheet = CreateNewSpreadsheet(filePath))
        {
            var worksheet = spreadsheet.AddWorksheet("Data");
            worksheet.AddCell(1, 1, "Name");
            worksheet.AddCell(2, 1, "Score");
            worksheet.AddCell(1, 2, "Alice");
            worksheet.AddCell(2, 2, 10);
            worksheet.AddCell(1, 3, "Bob");
            worksheet.AddCell(2, 3, 20);
            worksheet.AddCellWithFormula(2, 4, "SUM(B2:B3)");
        }

        OpenXmlValidation.AssertValid(filePath);

        using var reopened = OpenExistingSpreadsheet(filePath);
        var sheet = reopened.GetWorksheet("Data");

        Assert.NotNull(sheet);
        Assert.Equal("Name", sheet.GetCell(1, 1)?.GetStringValue());
        Assert.Equal("Alice", sheet.GetCell(1, 2)?.GetStringValue());
        Assert.Equal(20, sheet.GetCell(2, 3)?.GetIntValue());
        Assert.Equal("SUM(B2:B3)", sheet.GetCell(2, 4)?.GetFormula());
    }

    [Fact]
    public void WorkbookRoundtrip_RangeSortWithHeader_PreservesExpectedOrder()
    {
        var filePath = GetFilepath("reader-range-sort.xlsx");

        using (var spreadsheet = CreateNewSpreadsheet(filePath))
        {
            var worksheet = spreadsheet.AddWorksheet("Sheet1");
            var range = worksheet.GetRange("A1:B4");
            range.SetValues(
            [
                ["Name", "Score"],
                ["Alice", 10],
                ["Bob", 30],
                ["Cara", 20]
            ]);

            range.SortByColumn(2, SortDirection.Descending, hasHeader: true);
        }

        OpenXmlValidation.AssertValid(filePath);

        using var reopened = OpenExistingSpreadsheet(filePath);
        var sheet = reopened.GetWorksheet("Sheet1");

        Assert.NotNull(sheet);
        Assert.Equal("Name", sheet.GetCell(1, 1)?.GetStringValue());
        Assert.Equal("Bob", sheet.GetCell(1, 2)?.GetStringValue());
        Assert.Equal("Cara", sheet.GetCell(1, 3)?.GetStringValue());
        Assert.Equal("Alice", sheet.GetCell(1, 4)?.GetStringValue());
    }

    [Fact]
    public void WorkbookRoundtrip_TableMetadata_CanBeReadAfterReopen()
    {
        var filePath = GetFilepath("reader-table-roundtrip.xlsx");

        using (var spreadsheet = CreateNewSpreadsheet(filePath))
        {
            var worksheet = spreadsheet.AddWorksheet("Sheet1");
            worksheet.AddRows([ ["Item", "Qty"], ["Apple", 5], ["Pear", 8] ]);
            spreadsheet.AddTable("Sheet1", worksheet.AddCellOnIndex(1, 1), worksheet.AddCellOnIndex(2, 3), ["Item", "Qty"]);
        }

        OpenXmlValidation.AssertValid(filePath);

        using var reopened = OpenExistingSpreadsheet(filePath);
        var table = reopened.GetTable("Sheet1", "Table1");
        var allTables = reopened.GetTables("Sheet1").ToList();

        Assert.NotNull(table);
        Assert.Equal("Sheet1", table.WorksheetName);
        Assert.Equal(2, table.ColumnCount);
        Assert.Equal(["Item", "Qty"], table.ColumnNames);
        Assert.Single(allTables);
    }

    [Fact]
    public void WorkbookRoundtrip_WorksheetLifecycleAndCellMetadata_AreReadableAfterReopen()
    {
        var filePath = GetFilepath("reader-worksheet-lifecycle.xlsx");

        using (var spreadsheet = CreateNewSpreadsheet(filePath))
        {
            var primary = spreadsheet.AddWorksheet("First");
            primary.AddCell(1, 1, "Docs");
            primary.GetCell(1, 1)?.SetHyperlink("https://example.com");
            primary.GetCell(1, 1)?.SetComment("Review this cell", "Tests");

            spreadsheet.AddWorksheet("Second");
            spreadsheet.AddWorksheet("Third");
            spreadsheet.RenameWorksheet("Second", "Renamed");
            spreadsheet.MoveWorksheet("Renamed", 1);
            spreadsheet.SetWorksheetHidden("First", true);
            spreadsheet.RemoveWorksheet("Third");
        }

        OpenXmlValidation.AssertValid(filePath);

        using var reopened = OpenExistingSpreadsheet(filePath);
        var names = reopened.GetWorksheetsName().ToArray();
        var first = reopened.GetWorksheet("First");
        var firstCell = first?.GetCell(1, 1);

        Assert.Equal(["Renamed", "First"], names);
        Assert.NotNull(first);
        Assert.True(first.IsHidden);
        Assert.NotNull(firstCell);
        Assert.Equal("https://example.com/", firstCell.GetHyperlink());
        Assert.Equal("Review this cell", firstCell.GetComment());
    }

    [Fact]
    public void WorkbookRoundtrip_ValidationAndConditionalFormatting_ArePersistedAfterReopen()
    {
        var filePath = GetFilepath("reader-validation-formatting-roundtrip.xlsx");

        using (var spreadsheet = CreateNewSpreadsheet(filePath))
        {
            var worksheet = spreadsheet.AddWorksheet("Sheet1");
            var style = spreadsheet.CreateStyle(fill: new Fill(Color.LightYellow));
            var range = worksheet.GetRange("A1:A3");

            range.SetValues(
            [
                ["A"],
                ["B"],
                ["C"]
            ]);
            range.AddValidation(DataValidationOptions.List(["A", "B", "C"]));
            range.AddConditionalFormatting(ConditionalFormattingOptions.ContainsText("A", style));
        }

        OpenXmlValidation.AssertValid(filePath);

        using (var reopened = OpenExistingSpreadsheet(filePath))
        {
            var worksheet = reopened.GetWorksheet("Sheet1");
            Assert.NotNull(worksheet);
            Assert.Equal("A", worksheet.GetCell(1, 1)?.GetStringValue());
            Assert.Equal("C", worksheet.GetCell(1, 3)?.GetStringValue());
        }

        using var document = SpreadsheetDocument.Open(filePath, false);
        var worksheetPart = GetWorksheetPart(document, "Sheet1");
        var worksheetElement = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet element was not found.");
        var dataValidations = worksheetElement.GetFirstChild<OpenXml.DataValidations>() ?? throw new InvalidOperationException("DataValidations element was not found.");
        Assert.NotNull(dataValidations);
        Assert.NotEmpty(worksheetElement.Elements<OpenXml.ConditionalFormatting>());
    }

    [Fact]
    public void WorkbookRoundtrip_NamedRangeAndProtection_ArePersistedAfterReopen()
    {
        var filePath = GetFilepath("reader-namedrange-protection-roundtrip.xlsx");

        using (var spreadsheet = CreateNewSpreadsheet(filePath))
        {
            var worksheet = spreadsheet.AddWorksheet("Sheet1");
            worksheet.AddCell(1, 1, "Seed");
            var range = worksheet.GetRange("A1:A1");

            spreadsheet.AddNamedRange("SeedRange", range);
            worksheet.Protect("secret");
            spreadsheet.ProtectWorkbook("secret");
        }

        OpenXmlValidation.AssertValid(filePath);

        using (var reopened = OpenExistingSpreadsheet(filePath))
        {
            var worksheet = reopened.GetWorksheet("Sheet1");
            Assert.NotNull(worksheet);
            Assert.Equal("Seed", worksheet.GetCell(1, 1)?.GetStringValue());
        }

        using var document = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = document.WorkbookPart;
        Assert.NotNull(workbookPart);
        var workbook = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook element was not found.");
        var worksheetPart = GetWorksheetPart(document, "Sheet1");
        var seedRange = workbook.DefinedNames?.Elements<OpenXml.DefinedName>().SingleOrDefault(name => name.Name?.Value == "SeedRange");

        Assert.NotNull(seedRange);
        Assert.NotNull(workbook.GetFirstChild<OpenXml.WorkbookProtection>());
        Assert.NotNull(worksheetPart.Worksheet.GetFirstChild<OpenXml.SheetProtection>());
    }

    [Fact]
    public void WorkbookRoundtrip_SparseLookupPaths_AreReadableAfterReopen()
    {
        var filePath = GetFilepath("reader-sparse-lookups-roundtrip.xlsx");

        using (var spreadsheet = CreateNewSpreadsheet(filePath))
        {
            var worksheet = spreadsheet.AddWorksheet("Lookup");
            worksheet.AddCell(2, 3, "B3 value");
            worksheet.AddCell(5, 3, "E3 value");
            worksheet.AddCell(27, 10, "AA10 value");
        }

        OpenXmlValidation.AssertValid(filePath);

        using var reopened = OpenExistingSpreadsheet(filePath);
        var lookupWorksheet = reopened.GetWorksheet("Lookup");

        Assert.NotNull(lookupWorksheet);
        Assert.Equal("B3 value", lookupWorksheet.GetCellByReference("b3")?.GetStringValue());
        Assert.Equal("E3 value", lookupWorksheet.GetRow(3)?.GetCell(5)?.GetStringValue());
        Assert.Equal("AA10 value", lookupWorksheet.GetCellByReference("AA10")?.GetStringValue());
    }

    [Fact]
    public void WorkbookRoundtrip_SparseRowAppend_BackfillsMissingCellsAfterReopen()
    {
        var filePath = GetFilepath("reader-sparse-row-append-roundtrip.xlsx");

        using (var spreadsheet = CreateNewSpreadsheet(filePath))
        {
            var worksheet = spreadsheet.AddWorksheet("Lookup");
            worksheet.AddCell(2, 3, "B3 value");
            worksheet.AddCell(5, 3, "E3 value");
        }

        OpenXmlValidation.AssertValid(filePath);

        using (var reopened = OpenExistingSpreadsheet(filePath))
        {
            var worksheet = reopened.GetWorksheet("Lookup");
            var row = worksheet?.GetRow(3);

            Assert.NotNull(worksheet);
            Assert.NotNull(row);

            var appendedCell = row.AddCell(6, "F3 value");

            Assert.NotNull(row.GetCell(1));
            Assert.Equal("B3 value", row.GetCell(2)?.GetStringValue());
            Assert.NotNull(row.GetCell(3));
            Assert.NotNull(row.GetCell(4));
            Assert.Equal("E3 value", row.GetCell(5)?.GetStringValue());
            Assert.Equal("F3 value", appendedCell.GetStringValue());
        }

        OpenXmlValidation.AssertValid(filePath);

        using var verified = OpenExistingSpreadsheet(filePath);
        var verifiedRow = verified.GetWorksheet("Lookup")?.GetRow(3);

        Assert.NotNull(verifiedRow);
        Assert.NotNull(verifiedRow.GetCell(1));
        Assert.Equal("B3 value", verifiedRow.GetCell(2)?.GetStringValue());
        Assert.NotNull(verifiedRow.GetCell(3));
        Assert.NotNull(verifiedRow.GetCell(4));
        Assert.Equal("E3 value", verifiedRow.GetCell(5)?.GetStringValue());
        Assert.Equal("F3 value", verifiedRow.GetCell(6)?.GetStringValue());
    }

    private static WorksheetPart GetWorksheetPart(SpreadsheetDocument document, string worksheetName)
    {
        var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart was not found.");
        var workbook = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook element was not found.");
        var sheets = workbook.Sheets?.Elements<OpenXml.Sheet>() ?? throw new InvalidOperationException("Workbook sheets were not found.");
        var sheet = sheets.SingleOrDefault(candidate => string.Equals(candidate.Name?.Value, worksheetName, StringComparison.Ordinal));
        if (sheet == null)
        {
            throw new InvalidOperationException($"Worksheet '{worksheetName}' was not found.");
        }

        var sheetId = sheet.Id?.Value ?? throw new InvalidOperationException($"Worksheet '{worksheetName}' does not have a valid relationship id.");
        return (WorksheetPart)workbookPart.GetPartById(sheetId);
    }
}