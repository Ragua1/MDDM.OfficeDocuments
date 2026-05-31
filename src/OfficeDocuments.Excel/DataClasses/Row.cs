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

    public ICell? AddCellOnRange(uint beginColumn, uint endColumn, IStyle? style = null)
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
            AddCellOnIndex(i, style);
        }

        var mergedCell = GetCell(beginColumn);
        if (mergedCell == null)
        {
            return null;
        }

        var fromCell = mergedCell.CellReference;
        var toCellCell = GetCell(endColumn);
        if (toCellCell == null)
        {
            return null;
        }
        var toCell = toCellCell.CellReference;

        // Create the merged cell and append it to the MergeCells collection.
        OwnerWorksheet.AppendMergeReference($"{fromCell}:{toCell}");

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

        var cell = GetCell(columnIndex) ?? CreateCell(columnIndex);

        style = Style?.CreateMergedStyle(style) ?? style;

        cell.AddStyle(style);

        return cell;
    }

    private ICell CreateCell(uint columnIndex)
    {
        for (var i = _maxContiguousCellIndex + 1; i <= columnIndex; i++) // backfill missing earlier cells while preserving ordering
        {
            if (!_cellsByColumnIndex.ContainsKey(i))
            {
                InsertCell(new Cell(Worksheet, i, RowIndex));
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
