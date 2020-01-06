using OfficeDocumentsApi.Excel.Interfaces;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocumentsApi.Excel.Factory
{
    public interface ICellFactory
    {
        ICell CreateCell(IWorksheet worksheet, uint column, uint row, IStyle cellStyle = null);
        ICell CreateCell(IWorksheet worksheet, string cellReference, IStyle cellStyle);
        ICell CreateCell(IWorksheet worksheet, OpenXmlSpreadsheet.Cell element);
    }
}