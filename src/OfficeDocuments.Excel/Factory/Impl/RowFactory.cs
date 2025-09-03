using OfficeDocuments.Excel.DataClasses;
using OfficeDocuments.Excel.Interfaces;

namespace OfficeDocuments.Excel.Factory.Impl;

public class RowFactory : IRowFactory
{
    public IRow CreateRow(IWorksheet worksheet, uint rowIndex, IStyle cellStyle = null)
    {
        return new Row(worksheet, rowIndex, cellStyle);
    }

    public IRow CreateRow(IWorksheet worksheet, DocumentFormat.OpenXml.Spreadsheet.Row element)
    {
        return new Row(worksheet, element);
    }
}