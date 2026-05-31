using System.Collections.Generic;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Extensions;
using OfficeDocuments.Excel.Interfaces;
using OfficeDocuments.Excel.Options;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.DataClasses;

internal sealed class Range : Base, IRange
{
    private readonly Worksheet _worksheet;

    public new IWorksheet Worksheet => _worksheet;
    public uint FromColumn { get; }
    public uint FromRow { get; }
    public uint ToColumn { get; }
    public uint ToRow { get; }
    public string StartReference => CellExtension.GetExcelCellReference(FromColumn, FromRow);
    public string EndReference => CellExtension.GetExcelCellReference(ToColumn, ToRow);
    public string Reference => $"{StartReference}:{EndReference}";
    public IReadOnlyList<IRow> Rows => Enumerable.Range(0, checked((int)(ToRow - FromRow + 1)))
        .Select(offset => _worksheet.AddRow(FromRow + (uint)offset))
        .ToList();
    public IReadOnlyList<ICell> Cells => EnumerateCells().ToList();

    internal Range(Worksheet worksheet, uint fromColumn, uint fromRow, uint toColumn, uint toRow)
        : base(null)
    {
        _worksheet = worksheet;
        FromColumn = fromColumn;
        FromRow = fromRow;
        ToColumn = toColumn;
        ToRow = toRow;
    }

    public ICell? GetCell(uint columnIndex, uint rowIndex)
    {
        if (columnIndex < FromColumn || columnIndex > ToColumn || rowIndex < FromRow || rowIndex > ToRow)
        {
            return null;
        }

        return _worksheet.GetCell(columnIndex, rowIndex);
    }

    public ICell? GetCell(string reference)
    {
        var (rowIndex, columnIndex) = reference.GetExcelCellIndex();
        return GetCell(columnIndex, rowIndex);
    }

    public IReadOnlyList<IReadOnlyList<string?>> GetValues()
    {
        var result = new List<IReadOnlyList<string?>>();

        for (var rowIndex = FromRow; rowIndex <= ToRow; rowIndex++)
        {
            var rowValues = new List<string?>();
            for (var columnIndex = FromColumn; columnIndex <= ToColumn; columnIndex++)
            {
                rowValues.Add(_worksheet.GetCell(columnIndex, rowIndex)?.GetStringValue());
            }

            result.Add(rowValues);
        }

        return result;
    }

    public void SetValues(IEnumerable<IEnumerable<object?>> values)
    {
        ArgumentNullException.ThrowIfNull(values);

        var rowIndex = FromRow;
        foreach (var rowValues in values)
        {
            if (rowIndex > ToRow)
            {
                throw new ArgumentException("The provided values exceed the range height.", nameof(values));
            }

            var columnIndex = FromColumn;
            if (rowValues != null)
            {
                foreach (var value in rowValues)
                {
                    if (columnIndex > ToColumn)
                    {
                        throw new ArgumentException("The provided values exceed the range width.", nameof(values));
                    }

                    var cell = _worksheet.AddCellOnIndex(columnIndex, rowIndex);
                    if (value != null)
                    {
                        cell.SetValue(value);
                    }

                    columnIndex++;
                }
            }

            rowIndex++;
        }
    }

    public void ApplyStyle(IStyle? style)
    {
        if (style == null)
        {
            return;
        }

        foreach (var cell in EnsureCells())
        {
            cell.AddStyle(style);
        }
    }

    public void Merge()
    {
        EnsureCells();
        _worksheet.AppendMergeReference(Reference);
    }

    public void ApplyAutoFilter() => _worksheet.SetAutoFilter(Reference);

    public void SortByColumn(uint relativeColumnIndex, SortDirection direction = SortDirection.Ascending, bool hasHeader = false)
    {
        if (relativeColumnIndex < 1 || relativeColumnIndex > ToColumn - FromColumn + 1)
        {
            throw new ArgumentException($"Column index '{relativeColumnIndex}' is outside of the range.", nameof(relativeColumnIndex));
        }

        var columnIndex = FromColumn + relativeColumnIndex - 1;
        var rowSnapshots = new List<RowSnapshot>();
        for (var rowIndex = FromRow; rowIndex <= ToRow; rowIndex++)
        {
            var cells = new List<OpenXml.Cell?>();
            for (var currentColumn = FromColumn; currentColumn <= ToColumn; currentColumn++)
            {
                cells.Add((_worksheet.GetCell(currentColumn, rowIndex) as Cell)?.CloneElement());
            }

            var sortCell = _worksheet.GetCell(columnIndex, rowIndex);
            rowSnapshots.Add(new RowSnapshot(
                rowIndex,
                GetSortValue(sortCell),
                cells));
        }

        var dataRows = hasHeader ? rowSnapshots.Skip(1).ToList() : rowSnapshots;
        var orderedRows = direction == SortDirection.Descending
            ? dataRows.OrderByDescending(snapshot => snapshot.SortValue, SortValueComparer.Instance).ToList()
            : dataRows.OrderBy(snapshot => snapshot.SortValue, SortValueComparer.Instance).ToList();

        if (hasHeader && rowSnapshots.Count > 0)
        {
            orderedRows.Insert(0, rowSnapshots[0]);
        }

        for (var rowOffset = 0; rowOffset < orderedRows.Count; rowOffset++)
        {
            var targetRowIndex = FromRow + (uint)rowOffset;
            var snapshot = orderedRows[rowOffset];

            for (var columnOffset = 0; columnOffset < snapshot.Cells.Count; columnOffset++)
            {
                var targetColumnIndex = FromColumn + (uint)columnOffset;
                var targetCell = (Cell)_worksheet.AddCellOnIndex(targetColumnIndex, targetRowIndex);
                targetCell.ReplaceFrom(snapshot.Cells[columnOffset]);
            }
        }
    }

    public void AddValidation(DataValidationOptions options) => _worksheet.AddDataValidation(Reference, options);

    public void AddConditionalFormatting(ConditionalFormattingOptions options) => _worksheet.AddConditionalFormatting(Reference, options);

    private IEnumerable<ICell> EnumerateCells()
    {
        for (var rowIndex = FromRow; rowIndex <= ToRow; rowIndex++)
        {
            for (var columnIndex = FromColumn; columnIndex <= ToColumn; columnIndex++)
            {
                var cell = _worksheet.GetCell(columnIndex, rowIndex);
                if (cell != null)
                {
                    yield return cell;
                }
            }
        }
    }

    private IReadOnlyList<ICell> EnsureCells()
    {
        var cells = new List<ICell>();
        for (var rowIndex = FromRow; rowIndex <= ToRow; rowIndex++)
        {
            for (var columnIndex = FromColumn; columnIndex <= ToColumn; columnIndex++)
            {
                cells.Add(_worksheet.AddCellOnIndex(columnIndex, rowIndex));
            }
        }

        return cells;
    }

    private static SortValue GetSortValue(ICell? cell)
    {
        if (cell == null || !cell.HasValue() && !cell.HasFormula())
        {
            return SortValue.Empty;
        }

        if (cell.HasFormula())
        {
            return new SortValue(cell.GetFormula());
        }

        if (cell.TryGetValue(out decimal decimalValue))
        {
            return new SortValue(decimalValue);
        }

        if (cell.TryGetValue(out DateTime dateValue))
        {
            return new SortValue(dateValue);
        }

        if (cell.TryGetValue(out bool boolValue))
        {
            return new SortValue(boolValue ? 1M : 0M);
        }

        return new SortValue(cell.GetStringValue());
    }

    private sealed record RowSnapshot(uint RowIndex, SortValue SortValue, IReadOnlyList<OpenXml.Cell?> Cells);

    private readonly record struct SortValue(decimal? Number, DateTime? Date, string? Text)
    {
        public static SortValue Empty => new(null, null, null);

        public SortValue(decimal number) : this(number, null, null)
        {
        }

        public SortValue(DateTime date) : this(null, date, null)
        {
        }

        public SortValue(string? text) : this(null, null, text)
        {
        }
    }

    private sealed class SortValueComparer : IComparer<SortValue>
    {
        public static SortValueComparer Instance { get; } = new();

        public int Compare(SortValue x, SortValue y)
        {
            if (x.Number.HasValue || y.Number.HasValue)
            {
                return Nullable.Compare(x.Number, y.Number);
            }

            if (x.Date.HasValue || y.Date.HasValue)
            {
                return Nullable.Compare(x.Date, y.Date);
            }

            if (x.Text == null && y.Text == null)
            {
                return 0;
            }

            if (x.Text == null)
            {
                return -1;
            }

            if (y.Text == null)
            {
                return 1;
            }

            return StringComparer.OrdinalIgnoreCase.Compare(x.Text, y.Text);
        }
    }
}
