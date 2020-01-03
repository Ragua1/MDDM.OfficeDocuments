using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeDocumentsApi.Excel.Interfaces;
using Alignment = OfficeDocumentsApi.Excel.Styles.Alignment;
using Border = OfficeDocumentsApi.Excel.Styles.Border;
using Fill = OfficeDocumentsApi.Excel.Styles.Fill;
using Font = OfficeDocumentsApi.Excel.Styles.Font;
using NumberingFormat = OfficeDocumentsApi.Excel.Styles.NumberingFormat;

namespace OfficeDocumentsApi.Excel.Factory.Impl
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
            return new OfficeDocumentsApi.Excel.DataClasses.Worksheet(spreadsheet, worksheetPart, sheetData, cellStyle);
        }
    }
}