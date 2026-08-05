using DocumentFormat.OpenXml;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.DataClasses;

/// <summary>
/// Centralizes insertion of worksheet-level elements so new children honor the CT_Worksheet
/// child order. Each insert method encodes which existing elements the new element must follow;
/// when none are present the element is inserted right after the sheet data.
/// </summary>
internal sealed class WorksheetElementOrderer(SpreadsheetLib.Worksheet worksheetElement, SpreadsheetLib.SheetData sheetData)
{
    public void InsertConditionalFormatting(OpenXmlElement element) =>
        InsertAfterFirstPresent(element, LastConditionalFormatting, FirstMergeCells, FirstAutoFilter);

    public void InsertDataValidations(OpenXmlElement element) =>
        InsertAfterFirstPresent(element, FirstDataValidations, LastConditionalFormatting, FirstMergeCells, FirstAutoFilter);

    public void InsertHyperlinks(OpenXmlElement element) =>
        InsertAfterFirstPresent(element, FirstDataValidations, LastConditionalFormatting, FirstMergeCells, FirstAutoFilter);

    private void InsertAfterFirstPresent(OpenXmlElement element, params Func<OpenXmlElement?>[] predecessors)
    {
        foreach (var findPredecessor in predecessors)
        {
            if (findPredecessor() is { } predecessor)
            {
                worksheetElement.InsertAfter(element, predecessor);
                return;
            }
        }

        worksheetElement.InsertAfter(element, sheetData);
    }

    private OpenXmlElement? LastConditionalFormatting() => worksheetElement.Elements<SpreadsheetLib.ConditionalFormatting>().LastOrDefault();
    private OpenXmlElement? FirstMergeCells() => worksheetElement.GetFirstChild<SpreadsheetLib.MergeCells>();
    private OpenXmlElement? FirstAutoFilter() => worksheetElement.GetFirstChild<SpreadsheetLib.AutoFilter>();
    private OpenXmlElement? FirstDataValidations() => worksheetElement.GetFirstChild<SpreadsheetLib.DataValidations>();
}
