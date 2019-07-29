using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlApi.Interfaces;

namespace OpenXmlApi.DataClasses
{
    public class Row : IRow
    {
        public DocumentFormat.OpenXml.Spreadsheet.Row Element => throw new NotImplementedException();

        public IList<ICell> Cells => throw new NotImplementedException();

        public uint RowIndex => throw new NotImplementedException();

        public ICell CurrentCell => throw new NotImplementedException();

        public IWorksheet Worksheet => throw new NotImplementedException();

        public IStyle Style => throw new NotImplementedException();

        public ICell AddCell(IStyle style = null)
        {
            throw new NotImplementedException();
        }

        public ICell AddCell(uint columnIndex, IStyle style = null)
        {
            throw new NotImplementedException();
        }

        public ICell AddCellOnRange(uint beginColumn, uint endColumn, IStyle style = null)
        {
            throw new NotImplementedException();
        }

        public ICell AddCellWithFormula(string formula, IStyle style = null)
        {
            throw new NotImplementedException();
        }

        public ICell AddCellWithFormula(uint columnIndex, string formula, IStyle style = null)
        {
            throw new NotImplementedException();
        }

        public ICell AddCellWithValue<T>(T value, IStyle style = null)
        {
            throw new NotImplementedException();
        }

        public ICell AddCellWithValue<T>(uint columnIndex, T value, IStyle style = null)
        {
            throw new NotImplementedException();
        }

        public IStyle AddStyle(params IStyle[] styles)
        {
            throw new NotImplementedException();
        }

        public ICell GetCell(uint columnIndex)
        {
            throw new NotImplementedException();
        }

        public ICell GetCell(string columnName)
        {
            throw new NotImplementedException();
        }
    }
}
