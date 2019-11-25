using System;
using System.Collections.Generic;
using System.Linq;
using OfficeDocumentsApi.Excel.Interfaces;

namespace OfficeDocumentsApi.Excel.DataClasses
{
    internal class Row : Base, IRow
    {
        private static readonly HashSet<char> ColumnNames = new HashSet<char>(
            new []{ 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' });

        public DocumentFormat.OpenXml.Spreadsheet.Row Element { get; }
        public IList<ICell> Cells { get; } = new List<ICell>();
        public ICell CurrentCell => Cells.FirstOrDefault(x => x.ColumnIndex == currentCellIndex);

        public uint RowIndex { get; }

        private uint NextCellIndex => currentCellIndex + 1;
        private uint currentCellIndex = 0;

        internal Row(IWorksheet worksheet, uint rowIndex, IStyle cellStyle = null)
            : base(worksheet, cellStyle)
        {
            RowIndex = rowIndex;
            Element = new DocumentFormat.OpenXml.Spreadsheet.Row
            {
                RowIndex = rowIndex
            };
        }
        internal Row(IWorksheet worksheet, DocumentFormat.OpenXml.Spreadsheet.Row element)
            : base(worksheet, element.StyleIndex ?? 0)
        {
            RowIndex = element.RowIndex;
            Element = element;

            foreach (var cellElement in element.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
            {
                var cell = new Cell(Worksheet, cellElement);
                Cells.Insert(0, cell);

                if (cell.ColumnIndex > currentCellIndex)
                {
                    currentCellIndex = cell.ColumnIndex;
                }
            }
        }

        public ICell AddCell(IStyle style = null)
        {
            return AddCell(NextCellIndex, style);
        }

        public ICell AddCell(uint columnIndex, IStyle style = null)
        {
            return GetOrCreateCell(columnIndex, style);
        }

        public ICell AddCellWithValue<T>(T value, IStyle style = null)
        {
            return AddCellWithValue(NextCellIndex, value, style);
        }

        public ICell AddCellWithValue<T>(uint columnIndex, T value, IStyle style = null)
        {
            var cell = GetOrCreateCell(columnIndex, style);

            cell.SetValue(value);

            return cell;
        }

        public ICell AddCellWithFormula(string formula, IStyle style = null)
        {
            return AddCellWithFormula(NextCellIndex, formula, style);
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
            Worksheet.MergeCells.Append(mergeCell);

            return mergedCell;
        }

        public ICell GetCell(uint columnIndex)
        {
            if (columnIndex < 1)
            {
                throw new ArgumentException($"Invalid argument column index '{columnIndex}'");
            }
            return Cells?.FirstOrDefault(c => c.ColumnIndex == columnIndex);
        }

        public ICell GetCell(string columnName)
        {
            uint columnIndex = 0;
            columnName = columnName.ToUpper();

            for (var i = 0; i < columnName.Length; i++)
            {
                var ch = columnName[i];
                if (!ColumnNames.Contains(ch))
                {
                    throw new ArgumentException($"Invalid argument column name '{columnName}'");
                }

                columnIndex += (uint)(i * ('Z' - 'A') + (ch - 'A' + 1));
            }

            return GetCell(columnIndex);
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

            style = Style?.CreateMergedStyle(style) ?? style;

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
                    cell = new Cell(Worksheet, i, RowIndex);

                    Cells.Add(cell);
                    Element.Append(cell.Element);
                }
            }
            /*
            cell = new Cell(this.Worksheet, columnIndex, this.RowIndex);

            this.Cells.Insert(0, cell);
            this.Element.Append(cell.Element);
            */
            if (columnIndex > currentCellIndex)
            {
                currentCellIndex = columnIndex;
            }

            return cell;
        }
    }
}