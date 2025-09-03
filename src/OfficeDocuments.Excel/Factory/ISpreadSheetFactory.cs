using OfficeDocuments.Excel.Interfaces;
using OpenXml= DocumentFormat.OpenXml;

namespace OfficeDocuments.Excel.Factory;

public interface ISpreadSheetFactory
{
    ISpreadsheet CreateSpreadsheet(Stream stream, bool createNew);
    ISpreadsheet CreateSpreadsheet(string filePath, bool createNew);

    IWorksheet CreateWorksheet(Spreadsheet spreadsheet, OpenXml.Packaging.WorksheetPart worksheetPart, OpenXml.Spreadsheet.SheetData sheetData, IStyle cellStyle = null);
}