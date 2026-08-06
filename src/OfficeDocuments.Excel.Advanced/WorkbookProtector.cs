using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.Advanced;

/// <summary>
/// Owns workbook-structure protection and the legacy Excel password hash.
/// </summary>
internal sealed class WorkbookProtector(WorkbookPart workbookPart)
{
    public void Protect(string? password)
    {
        var workbookProtection = workbookPart.Workbook?.GetFirstChild<SpreadsheetLib.WorkbookProtection>();
        if (workbookProtection == null)
        {
            // workbookProtection must precede sheets in CT_Workbook; appending would invalidate the file.
            workbookProtection = new WorkbookElementOrderer(workbookPart.Workbook!)
                .Insert(new SpreadsheetLib.WorkbookProtection());
        }

        workbookProtection.LockStructure = true;
        if (!string.IsNullOrEmpty(password))
        {
            workbookProtection.WorkbookPassword = ComputeProtectionPassword(password);
        }
    }

    /// <summary>
    /// Computes the legacy 16-bit Excel protection-password hash. This is weak by design and
    /// exists only for compatibility with the workbook/worksheet protection format.
    /// </summary>
    public static HexBinaryValue ComputeProtectionPassword(string password)
    {
        var hash = 0;
        for (var index = password.Length - 1; index >= 0; index--)
        {
            hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
            hash ^= password[index];
        }

        hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
        hash ^= password.Length;
        hash ^= 0xCE4B;

        return new HexBinaryValue(hash.ToString("X4", CultureInfo.InvariantCulture));
    }
}
