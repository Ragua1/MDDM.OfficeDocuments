using OfficeDocuments.Excel;
using OfficeDocuments.Excel.DataClasses;
using OfficeDocuments.Excel.Interfaces;

namespace OfficeDocuments.Excel.Advanced;

/// <summary>
/// Advanced workbook-level features — structured tables, named ranges, and workbook protection —
/// surfaced as extension methods over the core <see cref="ISpreadsheet"/>. They drive the same
/// internal workbook state as the core coordinator, so they require the built-in
/// <see cref="Spreadsheet"/> implementation.
/// </summary>
public static class SpreadsheetAdvancedExtensions
{
    /// <summary>
    /// Adds a table to the specified worksheet.
    /// </summary>
    /// <exception cref="ArgumentException">Thrown when the worksheet cannot be found or the table definition is invalid.</exception>
    /// <exception cref="ArgumentNullException">Thrown when required parameters are null.</exception>
    public static ITableInfo AddTable(this ISpreadsheet spreadsheet, string worksheetName, ICell startCell, ICell endCell, List<string> columnsName)
        => Tables(spreadsheet).AddTable(worksheetName, startCell, endCell, columnsName);

    /// <summary>
    /// Adds a table over an existing range with optional creation options.
    /// </summary>
    public static ITableInfo AddTable(this ISpreadsheet spreadsheet, IRange range, List<string> columnsName, TableCreateOptions? options = null)
        => Tables(spreadsheet).AddTable(range, columnsName, options);

    /// <summary>
    /// Finds a table by name on the specified worksheet.
    /// </summary>
    public static ITableInfo? GetTable(this ISpreadsheet spreadsheet, string worksheetName, string tableName)
        => Tables(spreadsheet).GetTable(worksheetName, tableName);

    /// <summary>
    /// Returns all tables defined on the specified worksheet.
    /// </summary>
    public static IEnumerable<ITableInfo> GetTables(this ISpreadsheet spreadsheet, string worksheetName)
        => Tables(spreadsheet).GetTables(worksheetName);

    /// <summary>
    /// Returns all tables defined across the entire workbook.
    /// </summary>
    public static IEnumerable<ITableInfo> GetTables(this ISpreadsheet spreadsheet)
        => Tables(spreadsheet).GetTables();

    /// <summary>
    /// Renames an existing table. The new name must be unique within the workbook.
    /// </summary>
    public static void RenameTable(this ISpreadsheet spreadsheet, string worksheetName, string tableName, string newName)
        => Tables(spreadsheet).RenameTable(worksheetName, tableName, newName);

    /// <summary>
    /// Resizes an existing table to cover a new range. The column count must remain the same.
    /// </summary>
    public static void ResizeTable(this ISpreadsheet spreadsheet, string worksheetName, string tableName, IRange newRange)
        => Tables(spreadsheet).ResizeTable(worksheetName, tableName, newRange);

    /// <summary>
    /// Removes a table from the specified worksheet.
    /// </summary>
    public static void RemoveTable(this ISpreadsheet spreadsheet, string worksheetName, string tableName)
        => Tables(spreadsheet).RemoveTable(worksheetName, tableName);

    /// <summary>
    /// Adds a named range.
    /// </summary>
    public static void AddNamedRange(this ISpreadsheet spreadsheet, string name, IRange range, bool worksheetScoped = false)
    {
        var core = Core(spreadsheet);
        new NamedRangeManager(core.WorkbookPartInternal, core.GetSheetIndexByWorksheetName).Add(name, range, worksheetScoped);
    }

    /// <summary>
    /// Protects workbook structure metadata.
    /// </summary>
    public static void ProtectWorkbook(this ISpreadsheet spreadsheet, string? password = null)
        => new WorkbookProtector(Core(spreadsheet).WorkbookPartInternal).Protect(password);

    private static TableManager Tables(ISpreadsheet spreadsheet)
    {
        var core = Core(spreadsheet);
        return new TableManager(core.WorkbookPartInternal, core.GetWorksheetOrThrow, () => core.WorksheetCatalog);
    }

    private static Spreadsheet Core(ISpreadsheet spreadsheet)
    {
        ArgumentNullException.ThrowIfNull(spreadsheet);
        return spreadsheet as Spreadsheet
            ?? throw new ArgumentException($"Advanced operations require the built-in {nameof(Spreadsheet)} implementation.", nameof(spreadsheet));
    }
}
