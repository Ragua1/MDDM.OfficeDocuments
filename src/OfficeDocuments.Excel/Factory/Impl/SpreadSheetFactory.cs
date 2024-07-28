using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeDocuments.Excel.Interfaces;
using Alignment = OfficeDocuments.Excel.Styles.Alignment;
using Border = OfficeDocuments.Excel.Styles.Border;
using Fill = OfficeDocuments.Excel.Styles.Fill;
using Font = OfficeDocuments.Excel.Styles.Font;
using NumberingFormat = OfficeDocuments.Excel.Styles.NumberingFormat;
using Worksheet = OfficeDocuments.Excel.DataClasses.Worksheet;

namespace OfficeDocuments.Excel.Factory.Impl
{
    public class SpreadSheetFactory : ISpreadSheetFactory
    {
        public ISpreadsheet CreateSpreadsheet(Stream stream, bool createNew)
        {
            return new Spreadsheet(stream, createNew);
        }

        public ISpreadsheet CreateSpreadsheet(string filePath, bool createNew)
        {
            return new Spreadsheet(filePath, createNew);
        }

        public IWorksheet CreateWorksheet(Spreadsheet spreadsheet, WorksheetPart worksheetPart, SheetData sheetData,
            IStyle cellStyle = null)
        {
            return new Worksheet(spreadsheet, worksheetPart, sheetData, cellStyle);
        }
    }
}