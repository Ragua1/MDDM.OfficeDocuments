using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
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
    /// Adds a table to the specified worksheet
    /// </summary>
    /// <param name="worksheetName">The name of the worksheet</param>
    /// <param name="startCell">The starting cell of the table</param>
    /// <param name="endCell">The ending cell of the table</param>
    /// <param name="columnsName">The names of the table columns</param>
    /// <exception cref="ArgumentException">Thrown when worksheet cannot be found or table definition is invalid</exception>
    /// <exception cref="ArgumentNullException">Thrown when required parameters are null</exception>
    void AddTable(string worksheetName, ICell startCell, ICell endCell, List<string> columnsName);

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
    IStyle CreateStyle(Stylesheet stylesheet, Excel_Styles_Font? font = null, Excel_Styles_Fill? fill = null, Excel_Styles_Border? border = null, Excel_Styles_NumberingFormat? numberFormat = null, Excel_Styles_Alignment? alignment = null);
}