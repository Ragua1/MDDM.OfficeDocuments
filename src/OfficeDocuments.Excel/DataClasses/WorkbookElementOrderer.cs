using DocumentFormat.OpenXml;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.DataClasses;

/// <summary>
/// Centralizes insertion of workbook-level elements so new children honor the CT_Workbook child
/// order. Appending at the end happens to work on a workbook this library just created, but it
/// produces a schema-invalid file as soon as a later sibling already exists — which is the normal
/// case for a workbook opened from Excel, where <c>calcPr</c> is almost always present.
/// </summary>
internal sealed class WorkbookElementOrderer(SpreadsheetLib.Workbook workbook)
{
    /// <summary>
    /// The CT_Workbook child sequence (ECMA-376 Part 1, 18.2.27).
    /// </summary>
    private static readonly string[] ChildOrder =
    [
        "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets",
        "functionGroups", "externalReferences", "definedNames", "calcPr", "oleSize",
        "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing",
        "fileRecoveryPr", "webPublishObjects", "extLst"
    ];

    /// <summary>
    /// Inserts <paramref name="element"/> before the first child that must follow it, or appends
    /// it when no such child exists.
    /// </summary>
    public T Insert<T>(T element)
        where T : OpenXmlElement
    {
        ArgumentNullException.ThrowIfNull(element);

        var position = PositionOf(element.LocalName);
        var successor = workbook.ChildElements.FirstOrDefault(child => PositionOf(child.LocalName) > position);

        return successor is null
            ? workbook.AppendChild(element)
            : workbook.InsertBefore(element, successor);
    }

    private static int PositionOf(string localName)
    {
        var index = Array.IndexOf(ChildOrder, localName);
        return index < 0 ? int.MaxValue : index;
    }
}
