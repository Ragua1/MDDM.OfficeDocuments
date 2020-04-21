﻿using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeDocumentsApi.Excel.Interfaces;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocumentsApi.Excel.DataClasses
{
    internal class Worksheet : Base, IWorksheet
    {
        public Spreadsheet Spreadsheet { get; }
        public SpreadsheetLib.SheetData Element { get; }
        //public string Name => Element.LocalName;
        internal WorksheetPart WorksheetPart { get; }
        public IRow CurrentRow => GetRow();
        public ICell CurrentCell => CurrentRow?.CurrentCell;

        // collection of existed rows
        public IList<IRow> Rows { get; } = new List<IRow>();
        // collection of existed cells with custom width
        public IList<ICell> Cells { get { return Rows?.SelectMany(c => c.Cells).ToList(); } }

        public SpreadsheetLib.Columns Columns
        {
            get
            {
                if (columns == null)
                {
                    columns = new SpreadsheetLib.Columns();
                    WorksheetPart.Worksheet.InsertBefore(columns, WorksheetPart.Worksheet.Elements<SpreadsheetLib.SheetData>().First());
                }
                return columns;
            }
        }
        public SpreadsheetLib.MergeCells MergeCells
        {
            get
            {
                if (mergeCells == null)
                {
                    mergeCells = new SpreadsheetLib.MergeCells();
                    WorksheetPart.Worksheet.InsertAfter(mergeCells, WorksheetPart.Worksheet.Elements<SpreadsheetLib.SheetData>().First());
                }
                return mergeCells;
            }
        }

        private uint NextRowIndex => (CurrentRow?.RowIndex ?? 0) + 1;
        private uint NextCellIndex => (CurrentCell?.ColumnIndex ?? 0) + 1;

        private uint currentRow = 1;
        private SpreadsheetLib.Columns columns;
        private SpreadsheetLib.MergeCells mergeCells;

        internal Worksheet(Spreadsheet spreadsheet, WorksheetPart worksheetPart, SpreadsheetLib.SheetData sheetData, IStyle cellStyle = null) : base(null, cellStyle)
        {
            Spreadsheet = spreadsheet;
            WorksheetPart = worksheetPart;
            Element = sheetData;
            Worksheet = this;

            //load rows and cells
            var rows = sheetData.Elements<SpreadsheetLib.Row>();
            foreach (var rowElement in rows)
            {
                Rows.Add(new Row(this, rowElement));

                if (rowElement.RowIndex > currentRow)
                {
                    currentRow = rowElement.RowIndex;
                }
            }
        }

        public IRow AddRow(IStyle style = null)
        {
            return AddRow(NextRowIndex, style);
        }

        public IRow AddRow(uint rowIndex, IStyle style = null)
        {
            return GetOrCreateRow(rowIndex, style);
        }

        public ICell AddCell(IStyle style)
        {
            return AddCell(NextCellIndex, currentRow, style);
        }

        public ICell AddCell(uint columnIndex, IStyle style = null)
        {
            return AddCell(columnIndex, currentRow, style);
        }

        public ICell AddCell(uint columnIndex, uint rowIndex, IStyle style = null)
        {
            var row = AddRow(rowIndex);

            return row.AddCell(columnIndex, style);
        }

        public ICell AddCellWithValue<T>(T value, IStyle style = null)
        {
            return AddCellWithValue(NextCellIndex, currentRow, value, style);
        }

        public ICell AddCellWithValue<T>(uint columnIndex, T value, IStyle style = null)
        {
            return AddCellWithValue(columnIndex, currentRow, value, style);
        }

        public ICell AddCellWithValue<T>(uint columnIndex, uint rowIndex, T value, IStyle style = null)
        {
            var row = AddRow(rowIndex);

            return row.AddCellWithValue(columnIndex, value, style);
        }

        public ICell AddCellWithFormula(string formula, IStyle style = null)
        {
            return AddCellWithFormula(NextCellIndex, currentRow, formula, style);
        }

        public ICell AddCellWithFormula(uint columnIndex, string formula, IStyle style = null)
        {
            return AddCellWithFormula(columnIndex, currentRow, formula, style);
        }

        public ICell AddCellWithFormula(uint columnIndex, uint rowIndex, string formula, IStyle style = null)
        {
            var row = AddRow(rowIndex);

            return row.AddCellWithFormula(columnIndex, formula, style);
        }

        public ICell AddCellOnRange(uint beginColumn, uint endColumn, IStyle style = null)
        {
            return AddCellOnRange(beginColumn, endColumn, currentRow, style);
        }

        public ICell AddCellOnRange(uint beginColumn, uint endColumn, uint rowIndex, IStyle style = null)
        {
            return AddRow(rowIndex).AddCellOnRange(beginColumn, endColumn, style);
        }

        public ICell AddCellOnRange(uint beginColumn, uint endColumn, uint beginRow, uint endRow, IStyle style = null)
        {
            if (beginColumn < 1)
            {
                throw new ArgumentException($"Invalid argument column index '{beginColumn}'");
            }
            if (beginRow < 1)
            {
                throw new ArgumentException($"Invalid argument row index '{beginColumn}'");
            }

            if (beginColumn > endColumn || beginRow > endRow)
            {
                return null;
            }

            for (var i = beginRow; i <= endRow; i++)
            {
                var row = AddRow(i, style);
                for (var j = beginColumn; j <= endColumn; j++)
                {
                    row.AddCell(j, style);
                }
            }

            var mergedCell = GetCell(beginColumn, beginRow);
            var fromCell = mergedCell.CellReference;
            var toCell = GetCell(endColumn, endRow).CellReference;

            // Create the merged cell and append it to the MergeCells collection.
            var mergeCell = new SpreadsheetLib.MergeCell { Reference = $"{fromCell}:{toCell}" };
            Worksheet.MergeCells.Append(mergeCell);

            return mergedCell;
        }

        public ICell GetCell(uint columnIndex)
        {
            if (columnIndex < 1)
            {
                throw new ArgumentException($"Invalid argument column index '{columnIndex}'");
            }
            return GetRow()?.GetCell(columnIndex);
        }

        public ICell GetCell(uint columnIndex, uint rowIndex)
        {
            if (columnIndex < 1)
            {
                throw new ArgumentException($"Invalid argument column index '{columnIndex}'");
            }
            if (rowIndex < 1)
            {
                throw new ArgumentException($"Invalid argument column index '{rowIndex}'");
            }
            return GetRow(rowIndex)?.GetCell(columnIndex);
        }

        public IRow GetRow()
        {
            return GetRow(currentRow);
        }

        public IRow GetRow(uint rowIndex)
        {
            if (rowIndex < 1)
            {
                throw new ArgumentException($"Invalid argument column index '{rowIndex}'");
            }
            return Rows?.FirstOrDefault(r => r.RowIndex == rowIndex);
        }

        private IRow GetOrCreateRow(uint rowIndex, IStyle style)
        {
            if (rowIndex < 1)
            {
                throw new ArgumentException($"Invalid argument row index '{rowIndex}'");
            }

            var row = GetRow(rowIndex);
            if (row == null)
            {
                row = new Row(this, rowIndex);

                style = Style?.CreateMergedStyle(style) ?? style;

                Rows.Insert(0, row);
                Element.Append(row.Element);

                if (rowIndex > currentRow)
                {
                    currentRow = rowIndex;
                }
            }

            row.AddStyle(style);

            return row;
        }

        public void SetColumnWidth(double widthValue)
        {
            SetColumnWidth(CurrentCell?.ColumnIndex ?? 0, widthValue);
        }

        public void SetColumnWidth(uint columnIndex, double widthValue)
        {
            if (columnIndex < 1 || widthValue < 0)
            {
                return;
            }

            var column = Columns.Elements<SpreadsheetLib.Column>().FirstOrDefault(c => c.Max == columnIndex);
            if (column == null)
            {
                column = new SpreadsheetLib.Column { BestFit = true, CustomWidth = false, Width = widthValue, Min = columnIndex, Max = columnIndex };
                Worksheet.Columns.Append(column);
            }
            else
            {
                column.Width = widthValue;
            }
        }
    }
}