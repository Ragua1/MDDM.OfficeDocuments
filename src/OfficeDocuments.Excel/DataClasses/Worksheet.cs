using System.Collections.Concurrent;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Excel.Extensions;
using OfficeDocuments.Excel.Interfaces;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.DataClasses;

internal partial class Worksheet : Base, IWorksheet
{
    private static readonly ConcurrentDictionary<Type, PropertyInfo[]> PropertyCache = new();
    private readonly Dictionary<uint, IRow> _rowsByIndex = [];
    private readonly Dictionary<string, ICell> _cellsByReference = new(StringComparer.OrdinalIgnoreCase);

    public Spreadsheet Spreadsheet { get; }
    public SpreadsheetLib.SheetData Element { get; }
    internal WorksheetPart WorksheetPart { get; }
    internal SpreadsheetLib.Worksheet WorksheetElement => WorksheetPart.Worksheet ?? throw new InvalidOperationException("The worksheet part does not contain a worksheet.");
    public string Name => Spreadsheet.GetWorksheetName(this);
    public bool IsHidden => Spreadsheet.IsWorksheetHidden(this);
    public IRow? CurrentRow => GetRow(_currentRow);
    public ICell? CurrentCell => CurrentRow?.CurrentCell;

    public List<IRow> Rows { get; } = [];
    public List<ICell> Cells => Rows.SelectMany(row => row.Cells).ToList();

    public SpreadsheetLib.Columns Columns
    {
        get
        {
            if (_columns == null)
            {
                _columns = WorksheetElement.GetFirstChild<SpreadsheetLib.Columns>();
                if (_columns == null)
                {
                    _columns = new SpreadsheetLib.Columns();
                    WorksheetElement.InsertBefore(_columns, Element);
                }
            }

            return _columns;
        }
    }

    public SpreadsheetLib.MergeCells MergeCells
    {
        get
        {
            if (_mergeCells == null)
            {
                _mergeCells = WorksheetElement.GetFirstChild<SpreadsheetLib.MergeCells>();
                if (_mergeCells == null)
                {
                    _mergeCells = new SpreadsheetLib.MergeCells();
                    WorksheetElement.InsertAfter(_mergeCells, Element);
                }
            }

            return _mergeCells;
        }
    }

    private uint NextRowIndex => (CurrentRow?.RowIndex ?? 0) + 1;
    private uint NextCellIndex => (CurrentCell?.ColumnIndex ?? 0) + 1;

    private uint _currentRow = 1;
    private SpreadsheetLib.Columns? _columns;
    private SpreadsheetLib.MergeCells? _mergeCells;

    internal Worksheet(Spreadsheet spreadsheet, WorksheetPart worksheetPart, SpreadsheetLib.SheetData sheetData, IStyle? cellStyle = null)
        : base(cellStyle)
    {
        Spreadsheet = spreadsheet;
        WorksheetPart = worksheetPart;
        Element = sheetData;
        Worksheet = this;

        _columns = WorksheetElement.GetFirstChild<SpreadsheetLib.Columns>();
        _mergeCells = WorksheetElement.GetFirstChild<SpreadsheetLib.MergeCells>();

        foreach (var rowElement in sheetData.Elements<SpreadsheetLib.Row>())
        {
            var row = new Row(this, rowElement);
            Rows.Add(row);
            RegisterRow(row);

            if ((rowElement.RowIndex ?? 0) > _currentRow)
            {
                _currentRow = rowElement.RowIndex!;
            }
        }
    }

    public IRow AddRow(IStyle? style = null) => AddRow(NextRowIndex, style);

    public IRow AddRow(uint rowIndex, IStyle? style = null) => GetOrCreateRow(rowIndex, style);

    public ICell AddCell(IStyle? style = null) => AddCellOnIndex(NextCellIndex, _currentRow, style);

    public ICell AddCell<T>(T value, IStyle? style = null) => AddCell(NextCellIndex, _currentRow, value, style);

    public ICell AddCell<T>(uint columnIndex, T value, IStyle? style = null) => AddCell(columnIndex, _currentRow, value, style);

    public ICell AddCellOnIndex(uint columnIndex, IStyle? style = null) => AddCell(columnIndex, _currentRow, style);

    public ICell AddCellOnIndex(uint columnIndex, uint rowIndex, IStyle? style = null)
    {
        var row = AddRow(rowIndex);
        return row.AddCellOnIndex(columnIndex, style);
    }

    public ICell AddCell<T>(uint columnIndex, uint rowIndex, T value, IStyle? style = null)
    {
        var row = AddRow(rowIndex);
        return row.AddCell(columnIndex, value, style);
    }

    [Obsolete("Use AddCell method instead")]
    public ICell AddCellWithValue<T>(T value, IStyle? style = null) => AddCellWithValue(NextCellIndex, _currentRow, value, style);

    [Obsolete("Use AddCell method instead")]
    public ICell AddCellWithValue<T>(uint columnIndex, T value, IStyle? style = null) => AddCellWithValue(columnIndex, _currentRow, value, style);

    [Obsolete("Use AddCell method instead")]
    public ICell AddCellWithValue<T>(uint columnIndex, uint rowIndex, T value, IStyle? style = null)
    {
        var row = AddRow(rowIndex);
        return row.AddCellWithValue(columnIndex, value, style);
    }

    public ICell AddCellWithFormula(string formula, IStyle? style = null) => AddCellWithFormula(NextCellIndex, _currentRow, formula, style);

    public ICell AddCellWithFormula(uint columnIndex, string formula, IStyle? style = null) => AddCellWithFormula(columnIndex, _currentRow, formula, style);

    public ICell AddCellWithFormula(uint columnIndex, uint rowIndex, string formula, IStyle? style = null)
    {
        var row = AddRow(rowIndex);
        return row.AddCellWithFormula(columnIndex, formula, style);
    }

    public ICell? AddCellOnRange(uint beginColumn, uint endColumn, IStyle? style = null) => AddCellOnRange(beginColumn, endColumn, _currentRow, style);

    public ICell? AddCellOnRange(uint beginColumn, uint endColumn, uint rowIndex, IStyle? style = null)
    {
        if (beginColumn < 1)
        {
            throw new ArgumentException($"Invalid argument column index '{beginColumn}'", nameof(beginColumn));
        }

        if (rowIndex < 1)
        {
            throw new ArgumentException($"Invalid argument row index '{rowIndex}'", nameof(rowIndex));
        }

        if (beginColumn >= endColumn)
        {
            return null;
        }

        var range = GetRange(beginColumn, rowIndex, endColumn, rowIndex);
        range.ApplyStyle(style);
        range.Merge();
        return GetCell(beginColumn, rowIndex);
    }

    public ICell? AddCellOnRange(uint beginColumn, uint endColumn, uint beginRow, uint endRow, IStyle? style = null)
    {
        if (beginColumn < 1)
        {
            throw new ArgumentException($"Invalid argument column index '{beginColumn}'", nameof(beginColumn));
        }

        if (beginRow < 1)
        {
            throw new ArgumentException($"Invalid argument row index '{beginRow}'", nameof(beginRow));
        }

        if (beginColumn > endColumn || beginRow > endRow)
        {
            return null;
        }

        var range = GetRange(beginColumn, beginRow, endColumn, endRow);
        range.ApplyStyle(style);
        range.Merge();
        return GetCell(beginColumn, beginRow);
    }

    public IRange GetRange(uint fromColumn, uint fromRow, uint toColumn, uint toRow)
    {
        if (!TryGetRange(fromColumn, fromRow, toColumn, toRow, out var range) || range == null)
        {
            throw new ArgumentException("Invalid range coordinates.");
        }

        return range;
    }

    public IRange GetRange(string reference)
    {
        if (!TryGetRange(reference, out var range) || range == null)
        {
            throw new ArgumentException($"Invalid range reference '{reference}'", nameof(reference));
        }

        return range;
    }

    public bool TryGetRange(uint fromColumn, uint fromRow, uint toColumn, uint toRow, out IRange? range)
    {
        range = null;

        if (fromColumn < 1 || toColumn < 1 || fromRow < 1 || toRow < 1 || fromColumn > toColumn || fromRow > toRow)
        {
            return false;
        }

        range = new Range(this, fromColumn, fromRow, toColumn, toRow);
        return true;
    }

    public bool TryGetRange(string reference, out IRange? range)
    {
        range = null;

        if (!reference.TryGetExcelRange(out var coordinates))
        {
            return false;
        }

        return TryGetRange(coordinates.fromColumn, coordinates.fromRow, coordinates.toColumn, coordinates.toRow, out range);
    }

    public ICell? GetCell(uint columnIndex)
    {
        if (columnIndex < 1)
        {
            throw new ArgumentException($"Invalid argument column index '{columnIndex}'", nameof(columnIndex));
        }

        return GetRow()?.GetCell(columnIndex);
    }

    public ICell? GetCell(uint columnIndex, uint rowIndex)
    {
        if (columnIndex < 1)
        {
            throw new ArgumentException($"Invalid argument column index '{columnIndex}'", nameof(columnIndex));
        }

        if (rowIndex < 1)
        {
            throw new ArgumentException($"Invalid argument row index '{rowIndex}'", nameof(rowIndex));
        }

        return GetRow(rowIndex)?.GetCell(columnIndex);
    }

    public ICell? GetCellByReference(string reference)
    {
        if (string.IsNullOrWhiteSpace(reference))
        {
            return null;
        }

        return _cellsByReference.TryGetValue(reference.Trim(), out var cell) ? cell : null;
    }

    public IRow? GetRow() => GetRow(_currentRow);

    public IRow? GetRow(uint rowIndex)
    {
        if (rowIndex < 1)
        {
            throw new ArgumentException($"Invalid argument row index '{rowIndex}'", nameof(rowIndex));
        }

        return _rowsByIndex.TryGetValue(rowIndex, out var row) ? row : null;
    }

    public IRange? AddRows(IEnumerable<IEnumerable<object?>> rows, IStyle? style = null)
    {
        ArgumentNullException.ThrowIfNull(rows);

        var startRow = NextRowIndex;
        var currentRowIndex = startRow;
        uint maxColumn = 0;

        foreach (var rowValues in rows)
        {
            uint currentColumn = 1;
            var row = AddRow(currentRowIndex);

            if (rowValues != null)
            {
                foreach (var value in rowValues)
                {
                    var cell = row.AddCellOnIndex(currentColumn, style);
                    if (value != null)
                    {
                        cell.SetValue(value);
                    }

                    currentColumn++;
                }
            }

            maxColumn = Math.Max(maxColumn, currentColumn - 1);
            currentRowIndex++;
        }

        return maxColumn == 0 || currentRowIndex == startRow
            ? null
            : GetRange(1, startRow, maxColumn, currentRowIndex - 1);
    }

    public IRange? AddRows<T>(IEnumerable<T> items, bool includeHeader = false, IStyle? headerStyle = null, IStyle? rowStyle = null)
    {
        ArgumentNullException.ThrowIfNull(items);

        var type = typeof(T);
        var isScalar = IsScalarType(type);
        var properties = isScalar ? [] : GetReadableProperties(type);

        if (!isScalar && properties.Length == 0)
        {
            throw new ArgumentException($"Type '{type.Name}' does not expose readable public instance properties.");
        }

        var startRow = NextRowIndex;
        var currentRowIndex = startRow;
        uint columnCount = isScalar ? 1U : Convert.ToUInt32(properties.Length);
        var wroteAnything = false;

        if (includeHeader)
        {
            var headerRow = AddRow(currentRowIndex);
            if (isScalar)
            {
                headerRow.AddCellOnIndex(1, headerStyle).SetValue("Value");
            }
            else
            {
                for (var columnIndex = 0; columnIndex < properties.Length; columnIndex++)
                {
                    headerRow.AddCellOnIndex((uint)columnIndex + 1, headerStyle).SetValue(properties[columnIndex].Name);
                }
            }

            wroteAnything = true;
            currentRowIndex++;
        }

        foreach (var item in items)
        {
            var row = AddRow(currentRowIndex);

            if (isScalar)
            {
                var cell = row.AddCellOnIndex(1, rowStyle);
                if (item != null)
                {
                    cell.SetValue(item);
                }
            }
            else
            {
                for (var columnIndex = 0; columnIndex < properties.Length; columnIndex++)
                {
                    var cell = row.AddCellOnIndex((uint)columnIndex + 1, rowStyle);
                    var value = properties[columnIndex].GetValue(item);
                    if (value != null)
                    {
                        cell.SetValue(value);
                    }
                }
            }

            wroteAnything = true;
            currentRowIndex++;
        }

        return wroteAnything
            ? GetRange(1, startRow, columnCount, currentRowIndex - 1)
            : null;
    }

    public void SetColumnWidth(double widthValue) => SetColumnWidth(CurrentCell?.ColumnIndex ?? 0, widthValue);

    public void SetColumnWidth(uint columnIndex, double widthValue)
    {
        if (columnIndex < 1 || widthValue < 0)
        {
            return;
        }

        var column = Columns.Elements<SpreadsheetLib.Column>().FirstOrDefault(c => (uint)(c.Max ?? 0U) == columnIndex);
        if (column == null)
        {
            column = new SpreadsheetLib.Column
            {
                BestFit = true,
                CustomWidth = true,
                Width = widthValue,
                Min = columnIndex,
                Max = columnIndex
            };
            Columns.Append(column);
        }
        else
        {
            column.Width = widthValue;
            column.CustomWidth = true;
            column.BestFit = true;
        }
    }

    public void FreezePanes(uint frozenColumns, uint frozenRows)
    {
        if (frozenColumns == 0 && frozenRows == 0)
        {
            ClearFrozenPanes();
            return;
        }

        var sheetViews = WorksheetElement.GetFirstChild<SpreadsheetLib.SheetViews>();
        if (sheetViews == null)
        {
            sheetViews = WorksheetElement.InsertAt(new SpreadsheetLib.SheetViews(), 0);
        }

        var sheetView = sheetViews.GetFirstChild<SpreadsheetLib.SheetView>();
        if (sheetView == null)
        {
            sheetView = sheetViews.AppendChild(new SpreadsheetLib.SheetView { WorkbookViewId = 0U });
        }

        var pane = sheetView.GetFirstChild<SpreadsheetLib.Pane>();
        if (pane == null)
        {
            pane = sheetView.AppendChild(new SpreadsheetLib.Pane());
        }

        pane.State = SpreadsheetLib.PaneStateValues.Frozen;
        pane.HorizontalSplit = frozenColumns > 0 ? frozenColumns : null;
        pane.VerticalSplit = frozenRows > 0 ? frozenRows : null;
        pane.TopLeftCell = CellExtension.GetExcelCellReference(Math.Max(1, frozenColumns + 1), Math.Max(1, frozenRows + 1));
        pane.ActivePane = frozenColumns > 0 && frozenRows > 0
            ? SpreadsheetLib.PaneValues.BottomRight
            : frozenColumns > 0
                ? SpreadsheetLib.PaneValues.TopRight
                : SpreadsheetLib.PaneValues.BottomLeft;
    }

    public void ClearFrozenPanes()
    {
        var sheetView = WorksheetElement.GetFirstChild<SpreadsheetLib.SheetViews>()?.GetFirstChild<SpreadsheetLib.SheetView>();
        sheetView?.GetFirstChild<SpreadsheetLib.Pane>()?.Remove();
    }

    public void AutoFitColumns()
    {
        if (Rows.Count == 0)
        {
            return;
        }

        var maxColumn = Rows
            .SelectMany(row => row.Cells)
            .DefaultIfEmpty()
            .Max(cell => cell?.ColumnIndex ?? 0);

        if (maxColumn == 0)
        {
            return;
        }

        AutoFitColumns(GetRange(1, 1, maxColumn, _currentRow));
    }

    public void AutoFitColumns(IRange range)
    {
        ArgumentNullException.ThrowIfNull(range);

        for (var columnIndex = range.FromColumn; columnIndex <= range.ToColumn; columnIndex++)
        {
            var maxLength = 0;
            for (var rowIndex = range.FromRow; rowIndex <= range.ToRow; rowIndex++)
            {
                var text = GetDisplayText(GetCell(columnIndex, rowIndex));
                maxLength = Math.Max(maxLength, text.Length);
            }

            if (maxLength > 0)
            {
                SetColumnWidth(columnIndex, Math.Min(Math.Max(maxLength + 2, 8), 80));
            }
        }
    }

    public void Protect(string? password = null)
    {
        var protection = WorksheetElement.GetFirstChild<SpreadsheetLib.SheetProtection>();
        if (protection == null)
        {
            protection = new SpreadsheetLib.SheetProtection();
            WorksheetElement.InsertAfter(protection, Element);
        }

        protection.Sheet = true;
        protection.Objects = true;
        protection.Scenarios = true;

        if (!string.IsNullOrEmpty(password))
        {
            protection.Password = WorkbookProtector.ComputeProtectionPassword(password);
        }
    }

    internal void AppendMergeReference(string reference)
    {
        if (MergeCells.Elements<SpreadsheetLib.MergeCell>().Any(mergeCell => mergeCell.Reference?.Value == reference))
        {
            return;
        }

        MergeCells.Append(new SpreadsheetLib.MergeCell { Reference = reference });
    }

    internal void SetAutoFilter(string reference)
    {
        var autoFilter = WorksheetElement.GetFirstChild<SpreadsheetLib.AutoFilter>();
        if (autoFilter == null)
        {
            autoFilter = new SpreadsheetLib.AutoFilter();
            WorksheetElement.InsertAfter(autoFilter, Element);
        }

        autoFilter.Reference = reference;
    }

    private IRow GetOrCreateRow(uint rowIndex, IStyle? style = null)
    {
        if (rowIndex < 1)
        {
            throw new ArgumentException($"Invalid argument row index '{rowIndex}'", nameof(rowIndex));
        }

        var row = GetRow(rowIndex);
        if (row == null)
        {
            var createdRow = new Row(this, rowIndex);
            row = createdRow;

            var insertionIndex = Rows.FindIndex(existingRow => existingRow.RowIndex > rowIndex);
            if (insertionIndex < 0)
            {
                Rows.Add(row);
            }
            else
            {
                Rows.Insert(insertionIndex, row);
            }

            RegisterRow(row);

            var nextRowElement = Element.Elements<SpreadsheetLib.Row>().FirstOrDefault(existingRow => (existingRow.RowIndex ?? 0) > rowIndex);
            if (nextRowElement == null)
            {
                Element.Append(createdRow.RowElement);
            }
            else
            {
                Element.InsertBefore(createdRow.RowElement, nextRowElement);
            }
        }

        if (rowIndex > _currentRow)
        {
            _currentRow = rowIndex;
        }

        style = Style?.CreateMergedStyle(style) ?? style;
        row.AddStyle(style);
        return row;
    }

    internal void RegisterCell(ICell cell)
    {
        _cellsByReference[cell.CellReference] = cell;
    }

    private void RegisterRow(IRow row)
    {
        _rowsByIndex[row.RowIndex] = row;
    }

    private static bool IsScalarType(Type type)
    {
        var actualType = Nullable.GetUnderlyingType(type) ?? type;
        return actualType.IsPrimitive
               || actualType.IsEnum
               || actualType == typeof(string)
               || actualType == typeof(decimal)
               || actualType == typeof(DateTime)
               || actualType == typeof(Guid);
    }

    private static PropertyInfo[] GetReadableProperties(Type type)
    {
        return PropertyCache.GetOrAdd(type, static sourceType => sourceType
            .GetProperties(BindingFlags.Instance | BindingFlags.Public)
            .Where(property => property.CanRead && property.GetIndexParameters().Length == 0)
            .OrderBy(property => property.MetadataToken)
            .ToArray());
    }

    private static string GetDisplayText(ICell? cell)
    {
        if (cell == null)
        {
            return string.Empty;
        }

        if (cell.HasFormula())
        {
            return cell.GetFormula() ?? string.Empty;
        }

        return cell.GetStringValue() ?? string.Empty;
    }

}
