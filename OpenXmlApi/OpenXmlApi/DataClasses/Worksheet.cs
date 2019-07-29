using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlApi.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenXmlApi.DataClasses
{
    public class Worksheet : IWorksheet
    {
        public Spreadsheet Spreadsheet => throw new NotImplementedException();

        public SheetData Element => throw new NotImplementedException();

        public IRow CurrentRow => throw new NotImplementedException();

        public ICell CurrentCell => throw new NotImplementedException();

        public IList<IRow> Rows => throw new NotImplementedException();

        public IList<ICell> Cells => throw new NotImplementedException();

        public Columns Columns => throw new NotImplementedException();

        public MergeCells MergeCells => throw new NotImplementedException();

        public IStyle Style => throw new NotImplementedException();

        IWorksheet IBase.Worksheet => throw new NotImplementedException();

        public ICell AddCell(IStyle style = null)
        {
            throw new NotImplementedException();
        }

        public ICell AddCell(uint columnIndex, IStyle style = null)
        {
            throw new NotImplementedException();
        }

        public ICell AddCell(uint columnIndex, uint rowIndex, IStyle style = null)
        {
            throw new NotImplementedException();
        }

        public ICell AddCellOnRange(uint beginColumn, uint endColumn, IStyle style = null)
        {
            throw new NotImplementedException();
        }

        public ICell AddCellOnRange(uint beginColumn, uint endColumn, uint rowIndex, IStyle style = null)
        {
            throw new NotImplementedException();
        }

        public ICell AddCellOnRange(uint beginColumn, uint endColumn, uint beginRow, uint endRow, IStyle style = null)
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

        public ICell AddCellWithFormula(uint columnIndex, uint rowIndex, string formula, IStyle style = null)
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

        public ICell AddCellWithValue<T>(uint columnIndex, uint rowIndex, T value, IStyle style = null)
        {
            throw new NotImplementedException();
        }

        public IRow AddRow(IStyle style = null)
        {
            throw new NotImplementedException();
        }

        public IRow AddRow(uint rowIndex, IStyle style = null)
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

        public ICell GetCell(uint columnIndex, uint rowIndex)
        {
            throw new NotImplementedException();
        }

        public IRow GetRow()
        {
            throw new NotImplementedException();
        }

        public IRow GetRow(uint rowIndex)
        {
            throw new NotImplementedException();
        }

        public void SetColumnWidth(double widthValue)
        {
            throw new NotImplementedException();
        }

        public void SetColumnWidth(uint columnIndex, double widthValue)
        {
            throw new NotImplementedException();
        }
    }
}
