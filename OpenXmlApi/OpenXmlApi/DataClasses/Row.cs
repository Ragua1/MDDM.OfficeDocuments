using System;
using System.Collections.Generic;
using System.Linq;
using OpenXmlApi.Interfaces;

namespace OpenXmlApi
{
    internal class Row : Base, IRow
    {
        public DocumentFormat.OpenXml.Spreadsheet.Row Element { get; }
        public IList<ICell> Cells { get; } = new List<ICell>();
        public ICell CurrentCell => this.Cells.FirstOrDefault(x => x.ColumnIndex == this.currentCellIndex);

        public uint RowIndex { get; }

        private uint NextCellIndex => this.currentCellIndex + 1;
        private uint currentCellIndex = 0;

        internal Row(IWorksheet worksheet, uint rowIndex, IStyle cellStyle = null)
            : base(worksheet, cellStyle)
        {
            this.RowIndex = rowIndex;
            this.Element = new DocumentFormat.OpenXml.Spreadsheet.Row
            {
                RowIndex = rowIndex
            };
        }
        internal Row(IWorksheet worksheet, DocumentFormat.OpenXml.Spreadsheet.Row element)
            : base(worksheet, element.StyleIndex ?? 0)
        {
            this.RowIndex = element.RowIndex;
            this.Element = element;

            foreach (var cellElement in element.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
            {
                var cell = new Cell(this.Worksheet, cellElement);
                this.Cells.Insert(0, cell);

                if (cell.ColumnIndex > this.currentCellIndex)
                {
                    this.currentCellIndex = cell.ColumnIndex;
                }
            }
        }

        public ICell AddCell(IStyle style = null)
        {
            return AddCell(this.NextCellIndex, style);
        }

        public ICell AddCell(uint columnIndex, IStyle style = null)
        {
            return GetOrCreateCell(columnIndex, style);
        }

        public ICell AddCellWithValue<T>(T value, IStyle style = null)
        {
            return AddCellWithValue(this.NextCellIndex, value, style);
        }

        public ICell AddCellWithValue<T>(uint columnIndex, T value, IStyle style = null)
        {
            var cell = GetOrCreateCell(columnIndex, style);

            cell.SetValue(value);

            return cell;
        }

        public ICell AddCellWithFormula(string formula, IStyle style = null)
        {
            return AddCellWithFormula(this.NextCellIndex, formula, style);
        }

        public ICell AddCellWithFormula(uint columnIndex, string formula, IStyle style = null)
        {
            var cell = GetOrCreateCell(columnIndex, style);

            cell.SetFormula(formula);

            return cell;
        }

        public ICell AddCellOnRange(uint beginColumn, uint endColumn, IStyle style = null)
        {
            if (beginColumn < 1)
            {
                throw new ArgumentException($"Invalid argument column index '{beginColumn}'");
            }

            if (beginColumn >= endColumn)
            {
                return null;
            }

            for (var i = beginColumn; i <= endColumn; i++)
            {
                AddCell(i, style);
            }

            var mergedCell = GetCell(beginColumn);
            var fromCell = mergedCell.CellReference;
            var toCell = GetCell(endColumn).CellReference;

            // Create the merged cell and append it to the MergeCells collection.
            var mergeCell = new DocumentFormat.OpenXml.Spreadsheet.MergeCell { Reference = $"{fromCell}:{toCell}" };
            this.Worksheet.MergeCells.Append(mergeCell);

            return mergedCell;
        }

        public ICell GetCell(uint columnIndex)
        {
            if (columnIndex < 1)
            {
                throw new ArgumentException($"Invalid argument column index '{columnIndex}'");
            }
            return this.Cells?.FirstOrDefault(c => c.ColumnIndex == columnIndex);
        }

        public ICell GetCell(string columnName)
        {
            var columnNames = new[] { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };

            uint columnIndex = 0;
            columnName = columnName.ToUpper();

            for (var i = 0; i < columnName.Length; i++)
            {
                var ch = columnName[i];
                if (!columnNames.Contains(ch))
                {
                    throw new ArgumentException($"Invalid argument column name '{columnName}'");
                }

                columnIndex += (uint)(i * ('Z' - 'A') + (ch - 'A' + 1));
            }

            return this.GetCell(columnIndex);
        }

        private ICell GetOrCreateCell(uint columnIndex, IStyle style)
        {
            if (columnIndex < 1)
            {
                throw new ArgumentException($"Invalid argument column index '{columnIndex}'");
            }

            var cell = GetCell(columnIndex) ?? CreateCell(columnIndex);/*
            if (cell == null)
            {
                cell = new Cell(this.Worksheet, columnIndex, this.RowIndex);

                this.Cells.Insert(0, cell);
                this.Element.Append(cell.Element);

            }*/

            style = this.Style?.CreateMergedStyle(style) ?? style;

            cell.AddStyle(style);

            return cell;
        }

        private ICell CreateCell(uint columnIndex)
        {
            ICell cell = null;

            for (uint i = 1; i <= columnIndex; i++) // check if previous cells in same row exist
            {
                if (GetCell(i) == null) // add too missing previous cells in same row
                {
                    cell = new Cell(this.Worksheet, i, this.RowIndex);

                    this.Cells.Add(cell);
                    this.Element.Append(cell.Element);
                }
            }
            /*
            cell = new Cell(this.Worksheet, columnIndex, this.RowIndex);

            this.Cells.Insert(0, cell);
            this.Element.Append(cell.Element);
            */
            if (columnIndex > this.currentCellIndex)
            {
                this.currentCellIndex = columnIndex;
            }

            return cell;
        }
    }
}