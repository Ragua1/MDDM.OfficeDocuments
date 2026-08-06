using System.Collections.Generic;
using OfficeDocuments.Excel.Extensions;
using OfficeDocuments.Excel.Interfaces;

namespace OfficeDocuments.Excel.DataClasses;

internal class Row : Base, IRow
{
    private readonly Dictionary<uint, ICell> _cellsByColumnIndex = [];

    public DocumentFormat.OpenXml.Spreadsheet.Row Element { get; }
    internal DocumentFormat.OpenXml.Spreadsheet.Row RowElement => Element;
    public IList<ICell> Cells { get; } = new List<ICell>();
    public ICell? CurrentCell => _currentCellIndex == 0
        ? null
        : GetCell(_currentCellIndex) ?? CreateCell(_currentCellIndex);

    public uint RowIndex { get; }

    private uint NextCellIndex => _currentCellIndex + 1;
    private uint _currentCellIndex = 0;
    private uint _maxContiguousCellIndex = 0;

    internal Row(IWorksheet worksheet, uint rowIndex, IStyle? cellStyle = null)
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
        RowIndex = element.RowIndex ?? throw new InvalidOperationException();
        Element = element;

        foreach (var cellElement in element.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
        {
            var cell = new Cell(Worksheet, cellElement);
            Cells.Add(cell);
            RegisterCell(cell);

            if (cell.ColumnIndex > _currentCellIndex)
            {
                _currentCellIndex = cell.ColumnIndex;
            }
        }
    }

    public ICell AddCell(IStyle? style = null)
    {
        return AddCellOnIndex(NextCellIndex, style);
    }

    public ICell AddCell<T>(T value, IStyle? style = null)
    {
        return AddCell(NextCellIndex, value, style);
    }

    public ICell AddCellOnIndex(uint columnIndex, IStyle? style = null)
    {
        return GetOrCreateCell(columnIndex, style);
    }

    public ICell AddCell<T>(uint columnIndex, T value, IStyle? style = null)
    {
        var cell = GetOrCreateCell(columnIndex, style);

        cell.SetValue(value);

        return cell;
    }

    [Obsolete("Use AddCell method instead")]
    public ICell AddCellWithValue<T>(T value, IStyle? style = null)
    {
        return AddCellWithValue(NextCellIndex, value, style);
    }

    [Obsolete("Use AddCell method instead")]
    public ICell AddCellWithValue<T>(uint columnIndex, T value, IStyle? style = null)
    {
        var cell = GetOrCreateCell(columnIndex, style);

        cell.SetValue(value);

        return cell;
    }

    public ICell AddCellWithFormula(string formula, IStyle? style = null)
    {
        return AddCellWithFormula(NextCellIndex, formula, style);
    }

    public ICell AddCellWithFormula(uint columnIndex, string formula, IStyle? style = null)
    {
        var cell = GetOrCreateCell(columnIndex, style);

        cell.SetFormula(formula);

        return cell;
    }

    public ICell AddCellOnRange(uint beginColumn, uint endColumn, IStyle? style = null)
    {
        if (beginColumn < 1)
        {
            throw new ArgumentException($"Invalid argument column index '{beginColumn}'", nameof(beginColumn));
        }

        if (beginColumn > endColumn)
        {
            throw new ArgumentException($"End column '{endColumn}' is before begin column '{beginColumn}'", nameof(endColumn));
        }

        var mergedCell = AddCellOnIndex(beginColumn, style);
        var lastCell = mergedCell;
        for (var i = beginColumn + 1; i <= endColumn; i++)
        {
            lastCell = AddCellOnIndex(i, style);
        }

        // A range of one cell is not a merge — writing mergeCell ref="A1:A1" would put a degenerate
        // element into the file for a request that asked for nothing to be merged.
        if (beginColumn < endColumn)
        {
            OwnerWorksheet.AppendMergeReference($"{mergedCell.CellReference}:{lastCell.CellReference}");
        }

        return mergedCell;
    }

    public ICell? GetCell(string columnName)
    {
        return GetCell(columnName.GetExcelColumnIndex());
    }

    public ICell GetCellByReference(string reference) => Worksheet.GetCellByReference(reference)!;

    private ICell GetOrCreateCell(uint columnIndex, IStyle? style = null)
    {
        if (columnIndex < 1)
        {
            throw new ArgumentException($"Invalid argument column index '{columnIndex}'");
        }

        // The row style seeds a cell once, when the cell is created — CreateCell does that, for
        // the backfilled cells too. It must not be re-applied on a later access: it would stamp
        // the row's value back over a facet the caller has since set on the cell, so a font size
        // set on the cell would resolve to the row's size. That inverts the documented
        // sheet -> row -> cell precedence, and it only surfaces when something touches the same
        // cell twice — which Range.ApplyStyle followed by Range.Merge does.
        var cell = GetCell(columnIndex) ?? CreateCell(columnIndex);

        cell.AddStyle(style);

        return cell;
    }

    private ICell CreateCell(uint columnIndex)
    {
        for (var i = _maxContiguousCellIndex + 1; i <= columnIndex; i++) // backfill missing earlier cells while preserving ordering
        {
            if (!_cellsByColumnIndex.ContainsKey(i))
            {
                var backfilledCell = new Cell(Worksheet, i, RowIndex);

                // A backfilled cell belongs to the row and has to look like it. The workbook
                // default is skipped on purpose: an unstyled sheet still hands down a style whose
                // index is 0, and s="0" means exactly what leaving the attribute off means, so
                // applying it would only put a redundant attribute on every backfilled cell.
                if (Style is { StyleIndex: > 0 })
                {
                    backfilledCell.AddStyle(Style);
                }

                InsertCell(backfilledCell);
            }
        }

        if (columnIndex > _currentCellIndex)
        {
            _currentCellIndex = columnIndex;
        }

        return _cellsByColumnIndex[columnIndex];
    }

    private void InsertCell(Cell cell)
    {
        var insertionIndex = Cells.TakeWhile(existingCell => existingCell.ColumnIndex < cell.ColumnIndex).Count();
        if (insertionIndex >= Cells.Count)
        {
            Cells.Add(cell);
        }
        else
        {
            Cells.Insert(insertionIndex, cell);
        }

        RegisterCell(cell);

        var nextCell = Cells.Skip(insertionIndex + 1).FirstOrDefault() as Cell;

        if (nextCell == null)
        {
            RowElement.Append(cell.Element);
        }
        else
        {
            RowElement.InsertBefore(cell.Element, nextCell.Element);
        }
    }

    private void RegisterCell(ICell cell)
    {
        _cellsByColumnIndex[cell.ColumnIndex] = cell;
        UpdateContiguousCellIndex(cell.ColumnIndex);
        OwnerWorksheet.RegisterCell(cell);
    }

    private void UpdateContiguousCellIndex(uint columnIndex)
    {
        if (columnIndex != _maxContiguousCellIndex + 1)
        {
            return;
        }

        _maxContiguousCellIndex = columnIndex;
        while (_cellsByColumnIndex.ContainsKey(_maxContiguousCellIndex + 1))
        {
            _maxContiguousCellIndex++;
        }
    }

    public ICell? GetCell(uint columnIndex)
    {
        if (columnIndex < 1)
        {
            throw new ArgumentException($"Invalid argument column index '{columnIndex}'");
        }

        return _cellsByColumnIndex.TryGetValue(columnIndex, out var cell) ? cell : null;
    }
}
