using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using Alignment = OfficeDocumentsApi.Excel.Styles.Alignment;
using Border = OfficeDocumentsApi.Excel.Styles.Border;
using Fill = OfficeDocumentsApi.Excel.Styles.Fill;
using Font = OfficeDocumentsApi.Excel.Styles.Font;
using NumberingFormat = OfficeDocumentsApi.Excel.Styles.NumberingFormat;

namespace OfficeDocumentsApi.Excel.Interfaces
{
    public interface ISpreadsheet : IDisposable
    {
        /// <summary>
        /// Create worksheet and apply 'style'
        /// </summary>
        /// <param name="sheetName">Worksheet name</param>
        /// <param name="sheetStyle">Custom style for worksheet</param>
        /// <returns>Created worksheet</returns>
        IWorksheet AddWorksheet(string sheetName = null, IStyle sheetStyle = null);

        /// <summary>
        /// Create custom style
        /// </summary>
        /// <param name="font">Custom font styling</param>
        /// <param name="fill">Custom fill styling</param>
        /// <param name="border">Custom border styling</param>
        /// <param name="numberFormat">Custom number format styling</param>
        /// <param name="alignment">Custom alignment styling</param>
        /// <returns>Created style</returns>
        IStyle CreateStyle(Font font = null, Fill fill = null, Border border = null, NumberingFormat numberFormat = null, Alignment alignment = null);

        /// <summary>
        /// Get worksheet by name
        /// </summary>
        /// <param name="name"></param>
        /// <returns>Worksheet or null</returns>
        IWorksheet GetWorksheet(string name);

        void AddTable(string worksheetName, ICell startCell, ICell endCell, List<string> columnsName);

        IEnumerable<string> GetWorksheetsName();

        /// <summary>
        /// Save and close document
        /// </summary>
        void Close();

        IStyle CreateStyle(Stylesheet stylesheet, Font font = null, Fill fill = null, Border border = null, NumberingFormat numberFormat = null, Alignment alignment = null);
    }
}