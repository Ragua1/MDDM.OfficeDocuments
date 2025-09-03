using OfficeDocuments.Excel.Interfaces;
using OfficeDocuments.Excel.Styles;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.Factory;

public interface IStyleFactory
{
    IStyle CreateStyle(OpenXmlSpreadsheet.Stylesheet stylesheet, Font font = null, Fill fill = null, Border border = null, NumberingFormat numberFormat = null, Alignment alignment = null);
    IStyle CreateStyle(OpenXmlSpreadsheet.Stylesheet stylesheet, int fontId = 0, int fillId = 0, int borderId = 0, int numberFormatId = 0, Alignment alignment = null);
    IStyle CreateStyle(OpenXmlSpreadsheet.Stylesheet stylesheet, uint styleIndex);
}