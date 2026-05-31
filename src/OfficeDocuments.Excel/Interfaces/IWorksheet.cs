using System.ComponentModel;
using System.Collections.Generic;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Options;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.Interfaces;

/// <summary>
/// Interface of worksheet
/// </summary>
public interface IWorksheet : IBase
{
    /// <summary>
    /// Instance of Spreadsheet
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    [Obsolete("This property exposes the concrete Spreadsheet implementation. Prefer worksheet, range, and spreadsheet interface APIs.")]
    Spreadsheet Spreadsheet { get; }
    /// <summary>
    /// Instance of worksheet element
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    [Obsolete("This property exposes raw OpenXml sheet data. Prefer worksheet and range APIs.")]
    SpreadsheetLib.SheetData Element { get; }
    /// <summary>
    /// Worksheet name.
    /// </summary>
    string Name { get; }
    /// <summary>
    /// Indicates whether the worksheet is hidden.
    /// </summary>
    bool IsHidden { get; }
    /// <summary>
    /// Instance of row with highest 'RowIndex', or null when no rows exist
    /// </summary>
    IRow? CurrentRow { get; }
    /// <summary>
    /// Instance of cell with highest 'ColumnIndex' on current row, or null when no cells exist
    /// </summary>
    ICell? CurrentCell { get; }
    /// <summary>
    /// Collention of rows on sheet
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    List<IRow> Rows { get; }
    /// <summary>
    /// Collection of cells on sheet
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    List<ICell> Cells { get; }
    /// <summary>
    /// Instance of columns with custom width
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    [Obsolete("This property exposes raw OpenXml columns. Prefer SetColumnWidth(...) or AutoFitColumns(...).")]
    SpreadsheetLib.Columns Columns { get; }
    /// <summary>
    /// Instance of merged cells
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    [Obsolete("This property exposes raw OpenXml merge metadata. Prefer GetRange(...).Merge() or AddCellOnRange(...).")]
    SpreadsheetLib.MergeCells MergeCells { get; }

    /// <summary>
    /// Create row after current row and apply custom style.
    /// </summary>
    /// <param name="style">Custom style for row</param>
    /// <returns>Created row</returns>
    IRow AddRow(IStyle? style = null);

    /// <summary>
    /// Create or get row on 'rowIndex' and apply custom style.
    /// </summary>
    /// <param name="rowIndex">Index of row</param>
    /// <param name="style">Custom style for row</param>
    /// <returns>Row on 'rowIndex'</returns>
    IRow AddRow(uint rowIndex, IStyle? style = null);

    /// <summary>
    /// Create cell on current row after current cell and apply custom style.
    /// </summary>
    /// <param name="style">Custom style for cell</param>
    /// <returns>Created cell</returns>
    ICell AddCell(IStyle? style = null);

    /// <summary>
    /// Create or get cell on current row on 'columnIndex' and apply custom style.
    /// </summary>
    /// <param name="columnIndex">Index of column</param>
    /// <param name="style">Custom style for cell</param>
    /// <returns>Cell on current row on 'columnIndex'</returns>
    ICell AddCellOnIndex(uint columnIndex, IStyle? style = null);

    /// <summary>
    /// Create or get cell on 'rowIndex' on 'columnIndex' and apply custom style.
    /// </summary>
    /// <param name="columnIndex">Index of column</param>
    /// <param name="rowIndex">Index of row</param>
    /// <param name="style">Custom style for cell</param>
    /// <returns>Cell on 'rowIndex' on 'columnIndex'</returns>
    ICell AddCellOnIndex(uint columnIndex, uint rowIndex, IStyle? style = null);

    /// <summary>
    /// Create cell on current row after current cell, set 'value' and apply custom style.
    /// </summary>
    /// <param name="value">Cell value</param>
    /// <param name="style">Custom style for cell</param>
    /// <returns>Created cell</returns>
    ICell AddCell<T>(T value, IStyle? style = null);
        
    [Obsolete("Use AddCell method instead")]
    ICell AddCellWithValue<T>(T value, IStyle? style = null);

    /// <summary>
    /// Create or get cell on current row on 'columnIndex', set 'value' and apply custom style.
    /// </summary>
    /// <param name="columnIndex">Index of column</param>
    /// <param name="value">Cell value</param>
    /// <param name="style">Custom style for cell</param>
    /// <returns>Cell on current row on 'columnIndex'</returns>
    ICell AddCell<T>(uint columnIndex, T value, IStyle? style = null);
        
    [Obsolete("Use AddCell method instead")]
    ICell AddCellWithValue<T>(uint columnIndex, T value, IStyle? style = null);

    /// <summary>
    /// Create or get cell on 'rowIndex' on 'columnIndex', set 'value' and apply custom style.
    /// </summary>
    /// <param name="columnIndex">Index of column</param>
    /// <param name="rowIndex">Index of row</param>
    /// <param name="value">Cell value</param>
    /// <param name="style">Custom style for cell</param>
    /// <returns>Cell on 'rowIndex' on 'columnIndex'</returns>
    ICell AddCell<T>(uint columnIndex, uint rowIndex, T value, IStyle? style = null);
        
    [Obsolete("Use AddCell method instead")]
    ICell AddCellWithValue<T>(uint columnIndex, uint rowIndex, T value, IStyle? style = null);

    /// <summary>
    /// Create cell on current row after current cell, set 'formula' and apply custom style.
    /// </summary>
    /// <param name="formula">Cell formula</param>
    /// <param name="style">Custom style for cell</param>
    /// <returns>Created cell</returns>
    ICell AddCellWithFormula(string formula, IStyle? style = null);

    /// <summary>
    /// Create or get cell on current row on 'columnIndex', set 'formula' and apply custom style.
    /// </summary>
    /// <param name="columnIndex">Index of column</param>
    /// <param name="formula">Cell formula</param>
    /// <param name="style">Custom style for cell</param>
    /// <returns>Cell on current row on 'columnIndex'</returns>
    ICell AddCellWithFormula(uint columnIndex, string formula, IStyle? style = null);

    /// <summary>
    /// Create or get cell on 'rowIndex' on 'columnIndex', set 'formula' and apply custom style.
    /// </summary>
    /// <param name="columnIndex">Index of column</param>
    /// <param name="rowIndex">Index of row</param>
    /// <param name="formula">Cell formula</param>
    /// <param name="style">Custom style for cell</param>
    /// <returns>Cell on 'rowIndex' on 'columnIndex'</returns>
    ICell AddCellWithFormula(uint columnIndex, uint rowIndex, string formula, IStyle? style = null);

    /// <summary>
    /// Create and merge cells on current row from 'beginColumn' to 'endColumn'
    /// </summary>
    /// <param name="beginColumn">Begin column index</param>
    /// <param name="endColumn">End column index</param>
    /// <param name="style">Custom style for cell</param>
    /// <returns>Merged cell</returns>
    ICell? AddCellOnRange(uint beginColumn, uint endColumn, IStyle? style = null);

    /// <summary>
    /// Create and merge cells on 'rowIndex' row from 'beginColumn' to 'endColumn'
    /// </summary>
    /// <param name="rowIndex">Index of row</param>
    /// <param name="beginColumn">Begin column index</param>
    /// <param name="endColumn">End column index</param>
    /// <param name="style">Custom style for cell</param>
    /// <returns>Merged cell</returns>
    ICell? AddCellOnRange(uint beginColumn, uint endColumn, uint rowIndex, IStyle? style = null);

    /// <summary>
    /// Create and merge cells from 'beginReference' to 'endReference'
    /// </summary>
    /// <param name="beginColumn">Begin column index</param>
    /// <param name="endColumn">End column index</param>
    /// <param name="beginRow">Begin row index</param>
    /// <param name="endRow">End row index</param>
    /// <param name="style">Custom style for cell</param>
    /// <returns>Merged cell</returns>
    ICell? AddCellOnRange(uint beginColumn, uint endColumn, uint beginRow, uint endRow, IStyle? style = null);

    /// <summary>
    /// Gets a rectangular worksheet range.
    /// </summary>
    IRange GetRange(uint fromColumn, uint fromRow, uint toColumn, uint toRow);

    /// <summary>
    /// Gets a rectangular worksheet range from an A1 reference such as A1:C3.
    /// </summary>
    IRange GetRange(string reference);

    /// <summary>
    /// Tries to get a rectangular worksheet range.
    /// </summary>
    bool TryGetRange(uint fromColumn, uint fromRow, uint toColumn, uint toRow, out IRange? range);

    /// <summary>
    /// Tries to get a rectangular worksheet range from an A1 reference.
    /// </summary>
    bool TryGetRange(string reference, out IRange? range);

    /// <summary>
    /// Get cell on current row on 'columnIndex'
    /// </summary>
    /// <param name="columnIndex">Index of column</param>
    /// <returns>Cell on current row on 'columnIndex' or null</returns>
    ICell? GetCell(uint columnIndex);

    /// <summary>
    /// Get cell on 'rowIndex' on 'columnIndex'
    /// </summary>
    /// <param name="columnIndex">Index of column</param>
    /// <param name="rowIndex">Index of row</param>
    /// <returns>Cell on 'rowIndex' on 'columnIndex' or null</returns>
    ICell? GetCell(uint columnIndex, uint rowIndex);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="reference"></param>
    /// <returns></returns>
    ICell? GetCellByReference(string reference);

    /// <summary>
    /// Get current row
    /// </summary>
    /// <returns>Current row or null</returns>
    IRow? GetRow();

    /// <summary>
    /// Get row on 'rowIndex'
    /// </summary>
    /// <param name="rowIndex">Index of row</param>
    /// <returns>Row on 'rowIndex' or null</returns>
    IRow? GetRow(uint rowIndex);

    /// <summary>
    /// Adds rows from an enumerable of row values.
    /// </summary>
    IRange? AddRows(IEnumerable<IEnumerable<object?>> rows, IStyle? style = null);

    /// <summary>
    /// Adds rows from an enumerable of objects.
    /// </summary>
    IRange? AddRows<T>(IEnumerable<T> items, bool includeHeader = false, IStyle? headerStyle = null, IStyle? rowStyle = null);

    /// <summary>
    /// Set width of column for current cell
    /// </summary>
    /// <param name="widthValue">Width of column</param>
    void SetColumnWidth(double widthValue);

    /// <summary>
    /// Set width of column for 'columnIndex'
    /// </summary>
    /// <param name="columnIndex">Index of column</param>
    /// <param name="widthValue">Width of column</param>
    void SetColumnWidth(uint columnIndex, double widthValue);

    /// <summary>
    /// Freezes worksheet panes using counts of frozen columns and rows.
    /// </summary>
    void FreezePanes(uint frozenColumns, uint frozenRows);

    /// <summary>
    /// Clears any frozen panes.
    /// </summary>
    void ClearFrozenPanes();

    /// <summary>
    /// Auto fits all used worksheet columns.
    /// </summary>
    void AutoFitColumns();

    /// <summary>
    /// Auto fits columns in a specific range.
    /// </summary>
    void AutoFitColumns(IRange range);

    /// <summary>
    /// Protects the worksheet.
    /// </summary>
    void Protect(string? password = null);

    /// <summary>
    /// Embeds an image from a stream into the worksheet anchored across a rectangular range.
    /// Both column and row indexes are 1-based.
    /// The image stretches to fill the range defined by the anchor corners.
    /// </summary>
    /// <param name="imageStream">The image content stream.</param>
    /// <param name="imageType">The format of the image.</param>
    /// <param name="fromColumn">1-based start column of the anchor.</param>
    /// <param name="fromRow">1-based start row of the anchor.</param>
    /// <param name="toColumn">1-based end column of the anchor (inclusive right edge).</param>
    /// <param name="toRow">1-based end row of the anchor (inclusive bottom edge).</param>
    void AddImage(Stream imageStream, ImageType imageType, uint fromColumn, uint fromRow, uint toColumn, uint toRow);

    /// <summary>
    /// Embeds an image file into the worksheet anchored across a rectangular range.
    /// The image type is inferred from the file extension (.png, .jpg/.jpeg, .gif, .bmp, .tiff/.tif).
    /// Both column and row indexes are 1-based.
    /// </summary>
    /// <param name="filePath">Path to the image file.</param>
    /// <param name="fromColumn">1-based start column of the anchor.</param>
    /// <param name="fromRow">1-based start row of the anchor.</param>
    /// <param name="toColumn">1-based end column of the anchor (inclusive right edge).</param>
    /// <param name="toRow">1-based end row of the anchor (inclusive bottom edge).</param>
    void AddImage(string filePath, uint fromColumn, uint fromRow, uint toColumn, uint toRow);
}
