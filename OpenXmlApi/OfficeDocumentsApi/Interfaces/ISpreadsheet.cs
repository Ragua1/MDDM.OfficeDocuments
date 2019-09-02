using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeDocumentsApi.Interfaces;
using Alignment = OfficeDocumentsApi.Styles.Alignment;
using Border = OfficeDocumentsApi.Styles.Border;
using Fill = OfficeDocumentsApi.Styles.Fill;
using Font = OfficeDocumentsApi.Styles.Font;
using NumberingFormat = OfficeDocumentsApi.Styles.NumberingFormat;

namespace OfficeDocumentsApi.Interfaces
{
    public interface ISpreadsheet
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

        string[] GetWorksheetsName();

        /// <summary>
        /// Save and close document
        /// </summary>
        void Close();

        /// <summary>
        /// Close document resources
        /// </summary>
        void Dispose();

        IStyle CreateStyle(Stylesheet stylesheet, Font font = null, Fill fill = null, Border border = null, NumberingFormat numberFormat = null, Alignment alignment = null);
    }
}