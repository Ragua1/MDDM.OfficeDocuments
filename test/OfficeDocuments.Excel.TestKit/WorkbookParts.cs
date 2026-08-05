using DocumentFormat.OpenXml.Packaging;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.TestKit;

/// <summary>
/// Navigation helpers for asserting against the raw package, where the sheet name has to be
/// resolved to a relationship id before the part can be reached.
/// </summary>
public static class WorkbookParts
{
    /// <summary>
    /// Resolves the <see cref="WorksheetPart"/> backing the sheet named <paramref name="worksheetName"/>.
    /// </summary>
    public static WorksheetPart GetWorksheetPart(SpreadsheetDocument document, string worksheetName)
    {
        ArgumentNullException.ThrowIfNull(document);

        var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart was not found.");
        var workbook = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook element was not found.");
        var sheets = workbook.Sheets?.Elements<SpreadsheetLib.Sheet>() ?? throw new InvalidOperationException("Workbook sheets were not found.");
        var sheet = sheets.SingleOrDefault(candidate => string.Equals(candidate.Name?.Value, worksheetName, StringComparison.Ordinal))
                    ?? throw new InvalidOperationException($"Worksheet '{worksheetName}' was not found.");
        var relationshipId = sheet.Id?.Value
                             ?? throw new InvalidOperationException($"Worksheet '{worksheetName}' does not have a valid relationship id.");

        return (WorksheetPart)workbookPart.GetPartById(relationshipId);
    }

    /// <summary>
    /// The local names of the workbook's direct children, in document order — the shape that
    /// CT_Workbook ordering assertions are made against.
    /// </summary>
    public static List<string> WorkbookChildNames(SpreadsheetDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);

        var workbook = document.WorkbookPart?.Workbook ?? throw new InvalidOperationException("Workbook element was not found.");

        return workbook.ChildElements.Select(child => child.LocalName).ToList();
    }
}
