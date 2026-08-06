using OfficeDocuments.Excel.DataClasses;
using OfficeDocuments.Excel.Interfaces;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.Advanced;

/// <summary>
/// Advanced worksheet-level features — image embedding and worksheet protection — surfaced as
/// extension methods over the core <see cref="IWorksheet"/>. They require the built-in
/// <see cref="Worksheet"/> implementation.
/// </summary>
public static class WorksheetAdvancedExtensions
{
    /// <summary>
    /// Embeds an image from a stream into the worksheet anchored across a rectangular range.
    /// Both column and row indexes are 1-based. The image stretches to fill the anchor.
    /// </summary>
    public static void AddImage(this IWorksheet worksheet, Stream imageStream, ImageType imageType, uint fromColumn, uint fromRow, uint toColumn, uint toRow)
        => Image(worksheet).AddImage(imageStream, imageType, fromColumn, fromRow, toColumn, toRow);

    /// <summary>
    /// Embeds an image file into the worksheet anchored across a rectangular range.
    /// The image type is inferred from the file extension (.png, .jpg/.jpeg, .gif, .bmp, .tiff/.tif).
    /// Both column and row indexes are 1-based.
    /// </summary>
    public static void AddImage(this IWorksheet worksheet, string filePath, uint fromColumn, uint fromRow, uint toColumn, uint toRow)
        => Image(worksheet).AddImage(filePath, fromColumn, fromRow, toColumn, toRow);

    /// <summary>
    /// Protects the worksheet.
    /// </summary>
    public static void Protect(this IWorksheet worksheet, string? password = null)
    {
        var core = Core(worksheet);
        var worksheetElement = core.WorksheetElement;
        var protection = worksheetElement.GetFirstChild<SpreadsheetLib.SheetProtection>();
        if (protection == null)
        {
            protection = new SpreadsheetLib.SheetProtection();
            worksheetElement.InsertAfter(protection, core.Element);
        }

        protection.Sheet = true;
        protection.Objects = true;
        protection.Scenarios = true;

        if (!string.IsNullOrEmpty(password))
        {
            protection.Password = WorkbookProtector.ComputeProtectionPassword(password);
        }
    }

    private static WorksheetImageWriter Image(IWorksheet worksheet)
    {
        var core = Core(worksheet);
        return new WorksheetImageWriter(core.WorksheetPart, core.WorksheetElement);
    }

    private static Worksheet Core(IWorksheet worksheet)
    {
        ArgumentNullException.ThrowIfNull(worksheet);
        return worksheet as Worksheet
            ?? throw new ArgumentException($"Advanced operations require the built-in {nameof(Worksheet)} implementation.", nameof(worksheet));
    }
}
