using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlApi
{
    public interface IXNodeFactory
    {
        Workbook CreateWorkbook(int sheetCount);

    }

    public interface IDocumentFactory
    {
        ISpreadsheet CreateSpreadsheet(string filepath, bool open);
        ISpreadsheet CreateSpreadsheet(Stream stream, bool open);
    }

    internal class DocumentFactory : IDocumentFactory
    {
        public ISpreadsheet CreateSpreadsheet(string filepath, bool open)
        {
            return new Spreadsheet(filepath, open);
        }

        public ISpreadsheet CreateSpreadsheet(Stream stream, bool open)
        {
            return new Spreadsheet(stream, open);
        }
    }
}