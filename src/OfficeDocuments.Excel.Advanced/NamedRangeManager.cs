using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Excel.Extensions;
using OfficeDocuments.Excel.Interfaces;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.Advanced;

/// <summary>
/// Owns workbook defined names (named ranges). The sheet index for worksheet-scoped names is
/// resolved through the injected delegate to avoid coupling to the worksheet catalog.
/// </summary>
internal sealed class NamedRangeManager(WorkbookPart workbookPart, Func<string, int> getSheetIndexByWorksheetName)
{
    public void Add(string name, IRange range, bool worksheetScoped)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);
        ArgumentNullException.ThrowIfNull(range);

        if (!IsValidNamedRange(name))
        {
            throw new ArgumentException($"Named range '{name}' is not valid.", nameof(name));
        }

        // definedNames must precede calcPr in CT_Workbook, which an opened Excel workbook almost always has.
        var definedNames = workbookPart.Workbook?.DefinedNames
                           ?? new WorkbookElementOrderer(workbookPart.Workbook!).Insert(new SpreadsheetLib.DefinedNames());
        var localSheetId = worksheetScoped ? Convert.ToUInt32(getSheetIndexByWorksheetName(range.Worksheet.Name)) : (uint?)null;

        if (definedNames.Elements<SpreadsheetLib.DefinedName>().Any(definedName =>
                string.Equals(definedName.Name?.Value, name, StringComparison.OrdinalIgnoreCase)
                && (definedName.LocalSheetId == null && localSheetId == null || definedName.LocalSheetId?.Value == localSheetId)))
        {
            throw new ArgumentException($"Named range '{name}' already exists.", nameof(name));
        }

        definedNames.Append(new SpreadsheetLib.DefinedName
        {
            Name = name,
            LocalSheetId = localSheetId,
            Text = $"{range.Worksheet.Name}!{range.Reference}"
        });
    }

    private static bool IsValidNamedRange(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return false;
        }

        if (!char.IsLetter(name[0]) && name[0] != '_')
        {
            return false;
        }

        if (name.Any(character => !(char.IsLetterOrDigit(character) || character is '_' or '.')))
        {
            return false;
        }

        return !name.TryGetExcelRange(out _);
    }
}
