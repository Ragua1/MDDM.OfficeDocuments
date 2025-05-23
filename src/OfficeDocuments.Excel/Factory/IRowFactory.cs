﻿using OfficeDocuments.Excel.Interfaces;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.Factory
{
    public interface IRowFactory
    {
        IRow CreateRow(IWorksheet worksheet, uint rowIndex, IStyle cellStyle = null);
        IRow CreateRow(IWorksheet worksheet, OpenXmlSpreadsheet.Row element);
    }
}