using OfficeDocuments.Excel.Advanced;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.UnitTests;

/// <summary>
/// Regression cover for the CT_Workbook ordering bug found by the EXCEL-011 phase 1 validation
/// gate: <c>workbookProtection</c> was appended at the end of the workbook, but the schema
/// requires it before <c>sheets</c>. The <c>definedNames</c> case is the latent twin — appending
/// is only correct while nothing that must follow it (typically <c>calcPr</c>) exists.
/// </summary>
public class WorkbookElementOrdererTests
{
    private static string[] ChildNames(SpreadsheetLib.Workbook workbook) =>
        workbook.ChildElements.Select(child => child.LocalName).ToArray();

    [Fact]
    public void Insert_IntoEmptyWorkbook_Appends()
    {
        var workbook = new SpreadsheetLib.Workbook();

        new WorkbookElementOrderer(workbook).Insert(new SpreadsheetLib.WorkbookProtection());

        Assert.Equal(["workbookProtection"], ChildNames(workbook));
    }

    [Fact]
    public void Insert_WorkbookProtection_LandsBeforeSheets()
    {
        var workbook = new SpreadsheetLib.Workbook(new SpreadsheetLib.Sheets());

        new WorkbookElementOrderer(workbook).Insert(new SpreadsheetLib.WorkbookProtection());

        Assert.Equal(["workbookProtection", "sheets"], ChildNames(workbook));
    }

    [Fact]
    public void Insert_WorkbookProtection_LandsBetweenWorkbookPropertiesAndSheets()
    {
        var workbook = new SpreadsheetLib.Workbook(
            new SpreadsheetLib.WorkbookProperties(),
            new SpreadsheetLib.Sheets());

        new WorkbookElementOrderer(workbook).Insert(new SpreadsheetLib.WorkbookProtection());

        Assert.Equal(["workbookPr", "workbookProtection", "sheets"], ChildNames(workbook));
    }

    [Fact]
    public void Insert_DefinedNames_LandsBetweenSheetsAndCalculationProperties()
    {
        // The shape of a workbook opened from Excel, where calcPr is almost always present.
        var workbook = new SpreadsheetLib.Workbook(
            new SpreadsheetLib.Sheets(),
            new SpreadsheetLib.CalculationProperties());

        new WorkbookElementOrderer(workbook).Insert(new SpreadsheetLib.DefinedNames());

        Assert.Equal(["sheets", "definedNames", "calcPr"], ChildNames(workbook));
    }

    [Fact]
    public void Insert_DefinedNames_WithNothingAfterIt_Appends()
    {
        var workbook = new SpreadsheetLib.Workbook(new SpreadsheetLib.Sheets());

        new WorkbookElementOrderer(workbook).Insert(new SpreadsheetLib.DefinedNames());

        Assert.Equal(["sheets", "definedNames"], ChildNames(workbook));
    }

    [Fact]
    public void Insert_FileVersion_LandsFirst()
    {
        var workbook = new SpreadsheetLib.Workbook(
            new SpreadsheetLib.WorkbookProperties(),
            new SpreadsheetLib.Sheets());

        new WorkbookElementOrderer(workbook).Insert(new SpreadsheetLib.FileVersion());

        Assert.Equal(["fileVersion", "workbookPr", "sheets"], ChildNames(workbook));
    }

    [Fact]
    public void Insert_ReturnsTheInsertedInstance()
    {
        var workbook = new SpreadsheetLib.Workbook(new SpreadsheetLib.Sheets());
        var protection = new SpreadsheetLib.WorkbookProtection();

        var inserted = new WorkbookElementOrderer(workbook).Insert(protection);

        Assert.Same(protection, inserted);
        Assert.Same(protection, workbook.FirstChild);
    }

    [Fact]
    public void Insert_AllInReverseOrder_ProducesSchemaOrder()
    {
        var workbook = new SpreadsheetLib.Workbook(new SpreadsheetLib.Sheets());
        var orderer = new WorkbookElementOrderer(workbook);

        orderer.Insert(new SpreadsheetLib.CalculationProperties());
        orderer.Insert(new SpreadsheetLib.DefinedNames());
        orderer.Insert(new SpreadsheetLib.WorkbookProtection());
        orderer.Insert(new SpreadsheetLib.WorkbookProperties());

        Assert.Equal(
            ["workbookPr", "workbookProtection", "sheets", "definedNames", "calcPr"],
            ChildNames(workbook));
    }

    [Fact]
    public void Insert_NullElement_Throws()
    {
        var workbook = new SpreadsheetLib.Workbook();

        Assert.Throws<ArgumentNullException>(
            () => new WorkbookElementOrderer(workbook).Insert<SpreadsheetLib.DefinedNames>(null!));
    }
}
