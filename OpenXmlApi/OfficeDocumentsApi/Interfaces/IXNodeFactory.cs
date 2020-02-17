using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocumentsApi.Interfaces
{
    public interface IXNodeFactory
    {
        Workbook CreateWorkbook(int sheetCount);

    }

    public interface IDocumentFactory
    {
        ISpreadsheet CreateSpreadsheet(string filepath, bool open);
        ISpreadsheet CreateSpreadsheet(Stream stream, bool open);

        IWorksheet CreateWorkSheet(ISpreadsheet spreadsheet, uint column, uint row, IStyle cellStyle = null);
    }

    //internal class DocumentFactory : IDocumentFactory
    //{
    //    public ISpreadsheet CreateSpreadsheet(string filepath, bool open)
    //    {
    //        return new Spreadsheet(filepath, open);
    //    }

    //    public ISpreadsheet CreateSpreadsheet(Stream stream, bool open)
    //    {
    //        return new Spreadsheet(stream, open);
    //    }
    //}
}