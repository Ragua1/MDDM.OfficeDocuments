using System;
using System.Collections.Generic;

namespace OfficeDocuments.Excel.Interfaces
{
    /// <summary>
    /// Interface of row
    /// </summary>
    public interface IRow : IBase
    {
        /// <summary>
        /// Instance of Row element
        /// </summary>
        DocumentFormat.OpenXml.Spreadsheet.Row Element { get; }
        /// <summary>
        /// Collection of cells on row
        /// </summary>
        IList<ICell> Cells { get; }
        /// <summary>
        /// Index of row
        /// </summary>
        uint RowIndex { get; }
        /// <summary>
        /// Instance of cell with highest 'ColumnIndex' on current row
        /// </summary>
        ICell CurrentCell { get; }

        /// <summary>
        /// Create cell after current cell and apply custom style.
        /// </summary>
        /// <param name="style">Custom style for cell</param>
        /// <returns>Created cell</returns>
        ICell AddCell(IStyle? style = null);

        /// <summary>
        /// Create or get cell on 'columnIndex' and apply custom style.
        /// </summary>
        /// <param name="columnIndex">Index of column</param>
        /// <param name="style">Custom style for cell</param>
        /// <returns>Cell on 'columnIndex'</returns>
        ICell AddCell(uint columnIndex, IStyle? style = null);

        /// <summary>
        /// Create cell after current cell, set object value and apply custom style.
        /// </summary>
        /// <param name="value">Cell value</param>
        /// <param name="style">Custom style for cell</param>
        /// <returns>Created cell</returns>
        ICell AddCell<T>(T value, IStyle style);
        
        [Obsolete("Use AddCell method instead")]
        ICell AddCellWithValue<T>(T value, IStyle? style = null);

        /// <summary>
        /// Create or get cell on 'columnIndex', set 'value' and apply custom style.
        /// </summary>
        /// <param name="columnIndex">Index of column</param>
        /// <param name="value">Cell value</param>
        /// <param name="style">Custom style for cell</param>
        /// <returns>Cell on 'columnIndex'</returns>
        ICell AddCell<T>(uint columnIndex, T value, IStyle style = null);
        
        [Obsolete("Use AddCell method instead")]
        ICell AddCellWithValue<T>(uint columnIndex, T value, IStyle? style = null);

        /// <summary>
        /// Create cell after current cell, set 'formula' and apply custom style.
        /// </summary>
        /// <param name="formula">Cell formula</param>
        /// <param name="style">Custom style for cell</param>
        /// <returns>Created cell</returns>
        ICell AddCellWithFormula(string formula, IStyle? style = null);

        /// <summary>
        /// Create or get cell on 'columnIndex', set 'formula' and apply custom style.
        /// </summary>
        /// <param name="columnIndex">Index of column</param>
        /// <param name="formula">Cell formula</param>
        /// <param name="style">Custom style for cell</param>
        /// <returns>Cell on 'columnIndex'</returns>
        ICell AddCellWithFormula(uint columnIndex, string formula, IStyle? style = null);

        /// <summary>
        /// Create and merge cells from 'beginColumn' to 'endColumn'
        /// </summary>
        /// <param name="beginColumn">Begin column index</param>
        /// <param name="endColumn">End column index</param>
        /// <param name="style">Custom style for cell</param>
        /// <returns>Merged cell</returns>
        ICell AddCellOnRange(uint beginColumn, uint endColumn, IStyle? style = null);

        /// <summary>
        /// Get cell on 'columnIndex'
        /// </summary>
        /// <param name="columnIndex">Index of column</param>
        /// <returns>Cell on 'columnIndex' or null</returns>
        ICell? GetCell(uint columnIndex);

        ICell? GetCell(string columnName);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="reference"></param>
        /// <returns></returns>
        ICell GetCellByReference(string reference);
    }
}