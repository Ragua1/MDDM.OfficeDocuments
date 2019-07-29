using System.Collections.Generic;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlApi.Interfaces
{
    /// <summary>
    /// Interface of worksheet
    /// </summary>
    public interface IWorksheet : IBase
    {
        /// <summary>
        /// Instance of Spreadsheet
        /// </summary>
        Spreadsheet Spreadsheet { get; }
        /// <summary>
        /// Instance of worksheet element
        /// </summary>
        SpreadsheetLib.SheetData Element { get; }
        /// <summary>
        /// Instance of row with highest 'RowIndex'
        /// </summary>
        IRow CurrentRow { get; }
        /// <summary>
        /// Instance of cell with highest 'ColumnIndex' on current row
        /// </summary>
        ICell CurrentCell { get; }
        /// <summary>
        /// Collention of rows on sheet
        /// </summary>
        IList<IRow> Rows { get; }
        /// <summary>
        /// Collection of cells on sheet
        /// </summary>
        IList<ICell> Cells { get; }
        /// <summary>
        /// Instance of columns with custom width
        /// </summary>
        SpreadsheetLib.Columns Columns { get; }
        /// <summary>
        /// Instance of merged cells
        /// </summary>
        SpreadsheetLib.MergeCells MergeCells { get; }

        /// <summary>
        /// Create row after current row and apply custom style.
        /// </summary>
        /// <param name="style">Custom style for row</param>
        /// <returns>Created row</returns>
        IRow AddRow(IStyle style = null);

        /// <summary>
        /// Create or get row on 'rowIndex' and apply custom style.
        /// </summary>
        /// <param name="rowIndex">Index of row</param>
        /// <param name="style">Custom style for row</param>
        /// <returns>Row on 'rowIndex'</returns>
        IRow AddRow(uint rowIndex, IStyle style = null);

        /// <summary>
        /// Create cell on current row after current cell and apply custom style.
        /// </summary>
        /// <param name="style">Custom style for cell</param>
        /// <returns>Created cell</returns>
        ICell AddCell(IStyle style = null);

        /// <summary>
        /// Create or get cell on current row on 'columnIndex' and apply custom style.
        /// </summary>
        /// <param name="columnIndex">Index of column</param>
        /// <param name="style">Custom style for cell</param>
        /// <returns>Cell on current row on 'columnIndex'</returns>
        ICell AddCell(uint columnIndex, IStyle style = null);

        /// <summary>
        /// Create or get cell on 'rowIndex' on 'columnIndex' and apply custom style.
        /// </summary>
        /// <param name="columnIndex">Index of column</param>
        /// <param name="rowIndex">Index of row</param>
        /// <param name="style">Custom style for cell</param>
        /// <returns>Cell on 'rowIndex' on 'columnIndex'</returns>
        ICell AddCell(uint columnIndex, uint rowIndex, IStyle style = null);

        /// <summary>
        /// Create cell on current row after current cell, set 'value' and apply custom style.
        /// </summary>
        /// <param name="value">Cell value</param>
        /// <param name="style">Custom style for cell</param>
        /// <returns>Created cell</returns>
        ICell AddCellWithValue<T>(T value, IStyle style = null);

        /// <summary>
        /// Create or get cell on current row on 'columnIndex', set 'value' and apply custom style.
        /// </summary>
        /// <param name="columnIndex">Index of column</param>
        /// <param name="value">Cell value</param>
        /// <param name="style">Custom style for cell</param>
        /// <returns>Cell on current row on 'columnIndex'</returns>
        ICell AddCellWithValue<T>(uint columnIndex, T value, IStyle style = null);

        /// <summary>
        /// Create or get cell on 'rowIndex' on 'columnIndex', set 'value' and apply custom style.
        /// </summary>
        /// <param name="columnIndex">Index of column</param>
        /// <param name="rowIndex">Index of row</param>
        /// <param name="value">Cell value</param>
        /// <param name="style">Custom style for cell</param>
        /// <returns>Cell on 'rowIndex' on 'columnIndex'</returns>
        ICell AddCellWithValue<T>(uint columnIndex, uint rowIndex, T value, IStyle style = null);

        /// <summary>
        /// Create cell on current row after current cell, set 'formula' and apply custom style.
        /// </summary>
        /// <param name="formula">Cell formula</param>
        /// <param name="style">Custom style for cell</param>
        /// <returns>Created cell</returns>
        ICell AddCellWithFormula(string formula, IStyle style = null);

        /// <summary>
        /// Create or get cell on current row on 'columnIndex', set 'formula' and apply custom style.
        /// </summary>
        /// <param name="columnIndex">Index of column</param>
        /// <param name="formula">Cell formula</param>
        /// <param name="style">Custom style for cell</param>
        /// <returns>Cell on current row on 'columnIndex'</returns>
        ICell AddCellWithFormula(uint columnIndex, string formula, IStyle style = null);

        /// <summary>
        /// Create or get cell on 'rowIndex' on 'columnIndex', set 'formula' and apply custom style.
        /// </summary>
        /// <param name="columnIndex">Index of column</param>
        /// <param name="rowIndex">Index of row</param>
        /// <param name="formula">Cell formula</param>
        /// <param name="style">Custom style for cell</param>
        /// <returns>Cell on 'rowIndex' on 'columnIndex'</returns>
        ICell AddCellWithFormula(uint columnIndex, uint rowIndex, string formula, IStyle style = null);

        /// <summary>
        /// Create and merge cells on current row from 'beginColumn' to 'endColumn'
        /// </summary>
        /// <param name="beginColumn">Begin column index</param>
        /// <param name="endColumn">End column index</param>
        /// <param name="style">Custom style for cell</param>
        /// <returns>Merged cell</returns>
        ICell AddCellOnRange(uint beginColumn, uint endColumn, IStyle style = null);

        /// <summary>
        /// Create and merge cells on 'rowIndex' row from 'beginColumn' to 'endColumn'
        /// </summary>
        /// <param name="rowIndex">Index of row</param>
        /// <param name="beginColumn">Begin column index</param>
        /// <param name="endColumn">End column index</param>
        /// <param name="style">Custom style for cell</param>
        /// <returns>Merged cell</returns>
        ICell AddCellOnRange(uint beginColumn, uint endColumn, uint rowIndex, IStyle style = null);

        /// <summary>
        /// Create and merge cells from 'beginReference' to 'endReference'
        /// </summary>
        /// <param name="beginColumn">Begin column index</param>
        /// <param name="endColumn">End column index</param>
        /// <param name="beginRow">Begin row index</param>
        /// <param name="endRow">End row index</param>
        /// <param name="style">Custom style for cell</param>
        /// <returns>Merged cell</returns>
        ICell AddCellOnRange(uint beginColumn, uint endColumn, uint beginRow, uint endRow, IStyle style = null);

        /// <summary>
        /// Get cell on current row on 'columnIndex'
        /// </summary>
        /// <param name="columnIndex">Index of column</param>
        /// <returns>Cell on current row on 'columnIndex' or null</returns>
        ICell GetCell(uint columnIndex);

        /// <summary>
        /// Get cell on 'rowIndex' on 'columnIndex'
        /// </summary>
        /// <param name="columnIndex">Index of column</param>
        /// <param name="rowIndex">Index of row</param>
        /// <returns>Cell on 'rowIndex' on 'columnIndex' or null</returns>
        ICell GetCell(uint columnIndex, uint rowIndex);

        /// <summary>
        /// Get current row
        /// </summary>
        /// <returns>Current row or null</returns>
        IRow GetRow();

        /// <summary>
        /// Get row on 'rowIndex'
        /// </summary>
        /// <param name="rowIndex">Index of row</param>
        /// <returns>Row on 'rowIndex' or null</returns>
        IRow GetRow(uint rowIndex);

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
    }
}