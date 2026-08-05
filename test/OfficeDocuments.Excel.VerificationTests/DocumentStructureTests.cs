using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.TestKit;
using OfficeDocuments.Excel.TestKit.Validation;

namespace OfficeDocuments.Excel.VerificationTests;

/// <summary>
/// Document-level structure: schema child order and survival of a close/reopen cycle. These
/// assert on the package rather than on an API result, because the failures they guard against
/// are invisible to a round-trip through this library — the file still reads back correctly and
/// only Excel rejects it.
/// </summary>
public class DocumentStructureTests : SpreadsheetTestBase
{
    [Fact]
    public void ProtectWorkbookAndNamedRange_KeepWorkbookChildrenInSchemaOrder()
    {
        var filePath = GetFilepath("workbook-child-order.xlsx");

        using (var spreadsheet = CreateNewSpreadsheet(filePath))
        {
            var worksheet = spreadsheet.AddWorksheet("Sheet1");
            worksheet.AddCell(1, 1, "Seed");

            spreadsheet.AddNamedRange("SeedRange", worksheet.GetRange("A1:A1"));
            spreadsheet.ProtectWorkbook("secret");
        }

        OpenXmlValidation.AssertValid(filePath);

        using var document = SpreadsheetDocument.Open(filePath, false);
        var children = WorkbookParts.WorkbookChildNames(document);

        // CT_Workbook sequence: workbookProtection precedes sheets, which precedes definedNames.
        Assert.True(
            children.IndexOf("workbookProtection") < children.IndexOf("sheets"),
            $"workbookProtection must precede sheets, got: {string.Join(", ", children)}");
        Assert.True(
            children.IndexOf("sheets") < children.IndexOf("definedNames"),
            $"sheets must precede definedNames, got: {string.Join(", ", children)}");
    }

    [Fact]
    public void AddImage_DrawingElementAppearsBeforeTableParts()
    {
        var filePath = GetFilepath("image-order.xlsx");

        using (var spreadsheet = CreateNewSpreadsheet(filePath))
        {
            var worksheet = spreadsheet.AddWorksheet("Sheet1");
            worksheet.AddRows([["Header"], ["Data"]]);
            spreadsheet.AddTable("Sheet1", worksheet.AddCellOnIndex(1, 1), worksheet.AddCellOnIndex(1, 2), ["Header"]);

            using var stream = new MemoryStream(TestImages.MinimalPng());
            worksheet.AddImage(stream, ImageType.Png, 3, 1, 5, 3);
        }

        OpenXmlValidation.AssertValid(filePath);

        using var document = SpreadsheetDocument.Open(filePath, false);
        var worksheetElement = WorkbookParts.GetWorksheetPart(document, "Sheet1").Worksheet
                               ?? throw new InvalidOperationException("Worksheet element was not found.");
        var children = worksheetElement.ChildElements.ToList();
        var drawingIndex = children.FindIndex(child => child is Drawing);
        var tablePartsIndex = children.FindIndex(child => child is TableParts);

        Assert.True(drawingIndex >= 0, "Drawing element must exist");
        Assert.True(tablePartsIndex < 0 || drawingIndex < tablePartsIndex, "Drawing must appear before TableParts");
    }

    [Fact]
    public void HyperlinksCommentsNamedRangesAndProtection_PersistAfterReopen()
    {
        var filePath = GetFilepath("annotations-persistence.xlsx");

        using (var spreadsheet = CreateNewSpreadsheet(filePath))
        {
            var worksheet = spreadsheet.AddWorksheet("Sheet 1");
            var cell = worksheet.AddCell("Docs");
            cell.SetHyperlink("https://example.com");
            cell.SetComment("Review this cell", "Tests");

            spreadsheet.AddNamedRange("DocsCell", worksheet.GetRange("A1"));
            worksheet.Protect("secret");
            spreadsheet.ProtectWorkbook("secret");
        }

        OpenXmlValidation.AssertValid(filePath);

        using (var reopened = OpenExistingSpreadsheet(filePath))
        {
            var reopenedCell = reopened.GetWorksheet("Sheet 1")?.GetCell(1, 1);

            Assert.NotNull(reopenedCell);
            Assert.Equal("https://example.com/", reopenedCell.GetHyperlink());
            Assert.Equal("Review this cell", reopenedCell.GetComment());
        }

        using var document = SpreadsheetDocument.Open(filePath, false);
        var workbook = document.WorkbookPart?.Workbook ?? throw new InvalidOperationException("Workbook element was not found.");
        var worksheetPart = WorkbookParts.GetWorksheetPart(document, "Sheet 1");
        var definedName = workbook.DefinedNames?.Elements<DefinedName>().SingleOrDefault(name => name.Name?.Value == "DocsCell");

        Assert.NotNull(definedName);
        Assert.NotNull(worksheetPart.Worksheet.GetFirstChild<SheetProtection>());
        Assert.NotNull(workbook.GetFirstChild<WorkbookProtection>());
    }

    [Fact]
    public void ReopenAndSaveWithoutChanges_KeepsTheDocumentValid()
    {
        // Guards the family of defects where every open/save cycle duplicates a part or appends
        // an empty element. Also the precondition for any future golden-file comparison.
        var filePath = GetFilepath("resave-idempotence.xlsx");

        using (var spreadsheet = CreateNewSpreadsheet(filePath))
        {
            var worksheet = spreadsheet.AddWorksheet("Sheet1");
            worksheet.AddRows([["Item", "Qty"], ["Apple", 5]]);
        }

        var partCountAfterCreate = CountParts(filePath);

        using (var reopened = OpenExistingSpreadsheet(filePath))
        {
            reopened.Close();
        }

        OpenXmlValidation.AssertValid(filePath);
        Assert.Equal(partCountAfterCreate, CountParts(filePath));
    }

    private static int CountParts(string filePath)
    {
        using var document = SpreadsheetDocument.Open(filePath, false);

        return document.WorkbookPart?.Parts.Count() + 1 ?? 0;
    }
}
