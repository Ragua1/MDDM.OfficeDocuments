using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocumentsApi.Interfaces
{
    /// <summary>
    /// Interface of cell
    /// </summary>
    public interface ICell : IBase, IOpenXmlWrapper<Cell>
    {
        ///// <summary>
        ///// Instance of cell element
        ///// </summary>
        //DocumentFormat.OpenXml.Spreadsheet.Cell Element { get; }
        /// <summary>
        /// Excel reference of cell
        /// </summary>
        string CellReference { get; }
        /// <summary>
        /// Row index of cell
        /// </summary>
        uint RowIndex { get; }
        /// <summary>
        /// Column index of cell
        /// </summary>
        uint ColumnIndex { get; }
        string Value { get; set; }

        /// <summary>
        /// Set cell value
        /// </summary>
        /// <param name="value">Cell value</param>
        void SetValue(object value);

        /// <summary>
        /// Set cell boolean value
        /// </summary>
        /// <param name="value">Cell value</param>
        void SetValue(bool value);

        /// <summary>
        /// Set cell date value
        /// </summary>
        /// <param name="value">Cell value</param>
        void SetValue(DateTime value);

        /// <summary>
        /// Set cell string value
        /// </summary>
        /// <param name="value">Cell value</param>
        void SetValue(string value);

        /// <summary>
        /// Set cell formula
        /// </summary>
        /// <param name="formula">Cell formula</param>
        void SetFormula(string formula);

        /// <summary>
        /// Gets the formula.
        /// </summary>
        string GetFormula();

        /// <summary>
        /// Gets the value.
        /// </summary>
        string GetStringValue();
        /// <summary>
        /// Gets the bool value.
        /// </summary>
        bool GetBoolValue();
        /// <summary>
        /// Gets the int value.
        /// </summary>
        int GetIntValue();
        /// <summary>
        /// Gets the long value.
        /// </summary>
        long GetLongValue();
        /// <summary>
        /// Gets the double value.
        /// </summary>
        double GetDoubleValue();
        /// <summary>
        /// Gets the decimal value.
        /// </summary>
        decimal GetDecimalValue();
        /// <summary>
        /// Gets the date value.
        /// </summary>
        /// <param name="format">Date format</param>
        DateTime GetDateValue(string format = null);
        /// <summary>
        /// Tries the get bool value.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>Result of operation</returns>
        bool TryGetValue(out bool value);
        /// <summary>
        /// Tries the get int value.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>Result of operation</returns>
        bool TryGetValue(out int value);
        /// <summary>
        /// Tries the get long value.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>Result of operation</returns>
        bool TryGetValue(out long value);
        /// <summary>
        /// Tries the get double value.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>Result of operation</returns>
        bool TryGetValue(out double value);
        /// <summary>
        /// Tries the get decimal value.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>Result of operation</returns>
        bool TryGetValue(out decimal value);
        /// <summary>
        /// Tries the get string value.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>Result of operation</returns>
        bool TryGetValue(out string value);
        /// <summary>
        /// Tries the get date value.
        /// </summary>
        /// <param name="format">Date format</param>
        /// <param name="value">The value.</param>
        /// <returns>Result of operation</returns>
        bool TryGetValue(out DateTime value, string format = null);
        /// <summary>
        /// Determines whether this instance has value.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this cell has value; otherwise, <c>false</c>.
        /// </returns>
        bool HasValue();
        /// <summary>
        /// Determines whether this instance has formula.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this cell has formula; otherwise, <c>false</c>.
        /// </returns>
        bool HasFormula();
    }
}