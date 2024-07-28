using DocumentFormat.OpenXml.Spreadsheet;
using OfficeDocuments.Excel.Interfaces;

namespace OfficeDocuments.Excel.Factory.Impl
{
    public class CellFactory : ICellFactory
    {
        public ICell CreateCell(IWorksheet worksheet, uint column, uint row, IStyle cellStyle = null)
        {
            return new DataClasses.Cell(worksheet, column, row, cellStyle);
        }

        public ICell CreateCell(IWorksheet worksheet, string cellReference, IStyle cellStyle)
        {
            return new DataClasses.Cell(worksheet, cellReference, cellStyle);
        }

        public ICell CreateCell(IWorksheet worksheet, Cell element)
        {
            return new DataClasses.Cell(worksheet, element);
        }
    }
}