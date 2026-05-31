using System.Collections.Generic;
using System.ComponentModel;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeDocuments.Excel.Options;
using Excel_Styles_Alignment = OfficeDocuments.Excel.Styles.Alignment;
using Excel_Styles_Border = OfficeDocuments.Excel.Styles.Border;
using Excel_Styles_Fill = OfficeDocuments.Excel.Styles.Fill;
using Excel_Styles_Font = OfficeDocuments.Excel.Styles.Font;
using Excel_Styles_NumberingFormat = OfficeDocuments.Excel.Styles.NumberingFormat;
using Styles_Alignment = OfficeDocuments.Excel.Styles.Alignment;
using Styles_Border = OfficeDocuments.Excel.Styles.Border;
using Styles_Fill = OfficeDocuments.Excel.Styles.Fill;
using Styles_Font = OfficeDocuments.Excel.Styles.Font;
using Styles_NumberingFormat = OfficeDocuments.Excel.Styles.NumberingFormat;

namespace OfficeDocuments.Excel.Interfaces;

public interface ISpreadsheet : IDisposable
{
    /// <summary>
    /// Create worksheet and apply 'style'
    /// </summary>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="sheetStyle">Custom style for worksheet</param>
    /// <returns>Created worksheet</returns>
    IWorksheet AddWorksheet(string? sheetName = null, IStyle? sheetStyle = null);

    /// <summary>
    /// Create custom style
    /// </summary>
    /// <param name="font">Custom font styling</param>
    /// <param name="fill">Custom fill styling</param>
    /// <param name="border">Custom border styling</param>
    /// <param name="numberFormat">Custom number format styling</param>
    /// <param name="alignment">Custom alignment styling</param>
    /// <returns>Created style</returns>
    IStyle CreateStyle(Styles_Font? font = null, Styles_Fill? fill = null, Styles_Border? border = null, Styles_NumberingFormat? numberFormat = null, Styles_Alignment? alignment = null);

    /// <summary>
    /// Get worksheet by name
    /// </summary>
    /// <param name="name">The name of the worksheet to retrieve</param>
    /// <returns>Worksheet if found, null otherwise</returns>
    IWorksheet? GetWorksheet(string name);

    /// <summary>
    /// Renames a worksheet.
    /// </summary>
    void RenameWorksheet(string currentName, string newName);

    /// <summary>
    /// Removes a worksheet.
    /// </summary>
    void RemoveWorksheet(string name);

    /// <summary>
    /// Moves a worksheet to a new 1-based position.
    /// </summary>
    void MoveWorksheet(string name, uint newPosition);

    /// <summary>
    /// Copies a worksheet using basic worksheet content and metadata.
    /// </summary>
    IWorksheet CopyWorksheet(string sourceName, string? newName = null);

    /// <summary>
    /// Sets worksheet hidden state.
    /// </summary>
    void SetWorksheetHidden(string name, bool isHidden);

    /// <summary>
    /// Adds a table to the specified worksheet.
    /// </summary>
    /// <param name="worksheetName">The name of the worksheet.</param>
    /// <param name="startCell">The top-left cell of the table range.</param>
    /// <param name="endCell">The bottom-right cell of the table range.</param>
    /// <param name="columnsName">The ordered column header names.</param>
    /// <returns>Metadata describing the created table.</returns>
    /// <exception cref="ArgumentException">Thrown when the worksheet cannot be found or the table definition is invalid.</exception>
    /// <exception cref="ArgumentNullException">Thrown when required parameters are null.</exception>
    ITableInfo AddTable(string worksheetName, ICell startCell, ICell endCell, List<string> columnsName);

    /// <summary>
    /// Adds a table over an existing range with optional creation options.
    /// </summary>
    /// <param name="range">The range covered by the table.</param>
    /// <param name="columnsName">The ordered column header names.</param>
    /// <param name="options">Optional table creation options including name and style.</param>
    /// <returns>Metadata describing the created table.</returns>
    ITableInfo AddTable(IRange range, List<string> columnsName, TableCreateOptions? options = null);

    /// <summary>
    /// Finds a table by name on the specified worksheet.
    /// </summary>
    /// <param name="worksheetName">The name of the worksheet.</param>
    /// <param name="tableName">The table name to look up.</param>
    /// <returns>Table metadata if found; otherwise null.</returns>
    ITableInfo? GetTable(string worksheetName, string tableName);

    /// <summary>
    /// Returns all tables defined on the specified worksheet.
    /// </summary>
    IEnumerable<ITableInfo> GetTables(string worksheetName);

    /// <summary>
    /// Returns all tables defined across the entire workbook.
    /// </summary>
    IEnumerable<ITableInfo> GetTables();

    /// <summary>
    /// Renames an existing table.
    /// </summary>
    /// <param name="worksheetName">The name of the worksheet containing the table.</param>
    /// <param name="tableName">The current table name.</param>
    /// <param name="newName">The new table name. Must be unique within the workbook.</param>
    void RenameTable(string worksheetName, string tableName, string newName);

    /// <summary>
    /// Resizes an existing table to cover a new range.
    /// The column count of the new range must match the existing table columns.
    /// </summary>
    /// <param name="worksheetName">The name of the worksheet containing the table.</param>
    /// <param name="tableName">The table name to resize.</param>
    /// <param name="newRange">The new range. Column count must remain the same.</param>
    void ResizeTable(string worksheetName, string tableName, IRange newRange);

    /// <summary>
    /// Removes a table from the specified worksheet.
    /// </summary>
    /// <param name="worksheetName">The name of the worksheet containing the table.</param>
    /// <param name="tableName">The table name to remove.</param>
    void RemoveTable(string worksheetName, string tableName);

    /// <summary>
    /// Adds a named range.
    /// </summary>
    void AddNamedRange(string name, IRange range, bool worksheetScoped = false);

    /// <summary>
    /// Protects workbook structure metadata.
    /// </summary>
    void ProtectWorkbook(string? password = null);

    /// <summary>
    /// Gets the names of all worksheets in the document
    /// </summary>
    /// <returns>A collection of worksheet names</returns>
    IEnumerable<string> GetWorksheetsName();

    /// <summary>
    /// Save and close document
    /// </summary>
    void Close();

    /// <summary>
    /// Creates a style with the specified properties
    /// </summary>
    /// <param name="stylesheet">The stylesheet to use</param>
    /// <param name="font">Custom font styling</param>
    /// <param name="fill">Custom fill styling</param>
    /// <param name="border">Custom border styling</param>
    /// <param name="numberFormat">Custom number format styling</param>
    /// <param name="alignment">Custom alignment styling</param>
    /// <returns>The created style</returns>
    [EditorBrowsable(EditorBrowsableState.Never)]
    [Obsolete("This overload exposes raw OpenXml stylesheet plumbing. Prefer CreateStyle(...) without a Stylesheet parameter.")]
    IStyle CreateStyle(Stylesheet stylesheet, Excel_Styles_Font? font = null, Excel_Styles_Fill? fill = null, Excel_Styles_Border? border = null, Excel_Styles_NumberingFormat? numberFormat = null, Excel_Styles_Alignment? alignment = null);
}
