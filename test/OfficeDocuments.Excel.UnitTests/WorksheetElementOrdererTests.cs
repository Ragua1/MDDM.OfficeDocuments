using OfficeDocuments.Excel.DataClasses;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.UnitTests;

/// <summary>
/// The CT_Worksheet child sequence is what every advanced worksheet feature depends on, and
/// getting it wrong produces a file Excel repairs rather than an exception. The orderer is pure
/// DOM logic, so it can be pinned here without a package.
/// </summary>
public class WorksheetElementOrdererTests
{
    // CT_Worksheet order for the elements this type places:
    // sheetData < autoFilter < mergeCells < conditionalFormatting < dataValidations < hyperlinks
    private static (SpreadsheetLib.Worksheet Worksheet, WorksheetElementOrderer Orderer) CreateWorksheet()
    {
        var sheetData = new SpreadsheetLib.SheetData();
        var worksheet = new SpreadsheetLib.Worksheet(sheetData);

        return (worksheet, new WorksheetElementOrderer(worksheet, sheetData));
    }

    private static string[] ChildNames(SpreadsheetLib.Worksheet worksheet) =>
        worksheet.ChildElements.Select(child => child.LocalName).ToArray();

    [Fact]
    public void InsertConditionalFormatting_EmptyWorksheet_LandsAfterSheetData()
    {
        var (worksheet, orderer) = CreateWorksheet();

        orderer.InsertConditionalFormatting(new SpreadsheetLib.ConditionalFormatting());

        Assert.Equal(["sheetData", "conditionalFormatting"], ChildNames(worksheet));
    }

    [Fact]
    public void InsertConditionalFormatting_WithAutoFilter_LandsAfterAutoFilter()
    {
        var (worksheet, orderer) = CreateWorksheet();
        worksheet.Append(new SpreadsheetLib.AutoFilter());

        orderer.InsertConditionalFormatting(new SpreadsheetLib.ConditionalFormatting());

        Assert.Equal(["sheetData", "autoFilter", "conditionalFormatting"], ChildNames(worksheet));
    }

    [Fact]
    public void InsertConditionalFormatting_WithAutoFilterAndMergeCells_LandsAfterMergeCells()
    {
        var (worksheet, orderer) = CreateWorksheet();
        worksheet.Append(new SpreadsheetLib.AutoFilter());
        worksheet.Append(new SpreadsheetLib.MergeCells());

        orderer.InsertConditionalFormatting(new SpreadsheetLib.ConditionalFormatting());

        Assert.Equal(["sheetData", "autoFilter", "mergeCells", "conditionalFormatting"], ChildNames(worksheet));
    }

    [Fact]
    public void InsertConditionalFormatting_Repeated_AppendsAfterTheLastOne()
    {
        var (worksheet, orderer) = CreateWorksheet();
        var first = new SpreadsheetLib.ConditionalFormatting();
        var second = new SpreadsheetLib.ConditionalFormatting();

        orderer.InsertConditionalFormatting(first);
        orderer.InsertConditionalFormatting(second);

        Assert.Equal(["sheetData", "conditionalFormatting", "conditionalFormatting"], ChildNames(worksheet));
        Assert.Same(first, worksheet.ChildElements[1]);
        Assert.Same(second, worksheet.ChildElements[2]);
    }

    [Fact]
    public void InsertDataValidations_WithConditionalFormatting_LandsAfterIt()
    {
        var (worksheet, orderer) = CreateWorksheet();
        orderer.InsertConditionalFormatting(new SpreadsheetLib.ConditionalFormatting());

        orderer.InsertDataValidations(new SpreadsheetLib.DataValidations());

        Assert.Equal(["sheetData", "conditionalFormatting", "dataValidations"], ChildNames(worksheet));
    }

    [Fact]
    public void InsertDataValidations_EmptyWorksheet_LandsAfterSheetData()
    {
        var (worksheet, orderer) = CreateWorksheet();

        orderer.InsertDataValidations(new SpreadsheetLib.DataValidations());

        Assert.Equal(["sheetData", "dataValidations"], ChildNames(worksheet));
    }

    [Fact]
    public void InsertHyperlinks_WithDataValidations_LandsAfterIt()
    {
        var (worksheet, orderer) = CreateWorksheet();
        orderer.InsertDataValidations(new SpreadsheetLib.DataValidations());

        orderer.InsertHyperlinks(new SpreadsheetLib.Hyperlinks());

        Assert.Equal(["sheetData", "dataValidations", "hyperlinks"], ChildNames(worksheet));
    }

    [Fact]
    public void InsertAll_InAnyCallOrder_ProducesSchemaOrder()
    {
        var (worksheet, orderer) = CreateWorksheet();
        worksheet.Append(new SpreadsheetLib.MergeCells());

        // Deliberately the reverse of the schema order.
        orderer.InsertHyperlinks(new SpreadsheetLib.Hyperlinks());
        orderer.InsertDataValidations(new SpreadsheetLib.DataValidations());
        orderer.InsertConditionalFormatting(new SpreadsheetLib.ConditionalFormatting());

        Assert.Equal(
            ["sheetData", "mergeCells", "conditionalFormatting", "dataValidations", "hyperlinks"],
            ChildNames(worksheet));
    }
}
