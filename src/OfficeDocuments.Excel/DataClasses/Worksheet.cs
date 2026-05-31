using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Reflection;
using System.Text;
using Color = System.Drawing.Color;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Extensions;
using OfficeDocuments.Excel.Interfaces;
using OfficeDocuments.Excel.Options;
using Drw = DocumentFormat.OpenXml.Drawing;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;
using XdrSpr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeDocuments.Excel.DataClasses;

internal class Worksheet : Base, IWorksheet
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
            protection.Password = Spreadsheet.ComputeProtectionPassword(password);
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

    internal void AddDataValidation(string reference, DataValidationOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);

        var validations = WorksheetElement.GetFirstChild<SpreadsheetLib.DataValidations>();
        if (validations == null)
        {
            validations = new SpreadsheetLib.DataValidations();
            InsertAfterConditionalFormatting(validations);
        }

        var validation = new SpreadsheetLib.DataValidation
        {
            AllowBlank = options.AllowBlank,
            SequenceOfReferences = new ListValue<StringValue> { InnerText = reference }
        };

        validation.Type = options.Type switch
        {
            DataValidationType.List => SpreadsheetLib.DataValidationValues.List,
            DataValidationType.Whole => SpreadsheetLib.DataValidationValues.Whole,
            DataValidationType.Decimal => SpreadsheetLib.DataValidationValues.Decimal,
            DataValidationType.Date => SpreadsheetLib.DataValidationValues.Date,
            DataValidationType.Custom => SpreadsheetLib.DataValidationValues.Custom,
            _ => throw new ArgumentOutOfRangeException(nameof(options))
        };

        if (options.Operator.HasValue)
        {
            validation.Operator = options.Operator.Value switch
            {
                DataValidationOperator.Between => SpreadsheetLib.DataValidationOperatorValues.Between,
                DataValidationOperator.NotBetween => SpreadsheetLib.DataValidationOperatorValues.NotBetween,
                DataValidationOperator.Equal => SpreadsheetLib.DataValidationOperatorValues.Equal,
                DataValidationOperator.NotEqual => SpreadsheetLib.DataValidationOperatorValues.NotEqual,
                DataValidationOperator.GreaterThan => SpreadsheetLib.DataValidationOperatorValues.GreaterThan,
                DataValidationOperator.LessThan => SpreadsheetLib.DataValidationOperatorValues.LessThan,
                DataValidationOperator.GreaterThanOrEqual => SpreadsheetLib.DataValidationOperatorValues.GreaterThanOrEqual,
                DataValidationOperator.LessThanOrEqual => SpreadsheetLib.DataValidationOperatorValues.LessThanOrEqual,
                _ => throw new ArgumentOutOfRangeException(nameof(options))
            };
        }

        if (!string.IsNullOrWhiteSpace(options.PromptTitle))
        {
            validation.PromptTitle = options.PromptTitle;
        }

        if (!string.IsNullOrWhiteSpace(options.Prompt))
        {
            validation.Prompt = options.Prompt;
        }

        if (!string.IsNullOrWhiteSpace(options.ErrorTitle))
        {
            validation.ErrorTitle = options.ErrorTitle;
        }

        if (!string.IsNullOrWhiteSpace(options.Error))
        {
            validation.Error = options.Error;
        }

        validation.Append(new SpreadsheetLib.Formula1(options.Formula1));
        if (!string.IsNullOrWhiteSpace(options.Formula2))
        {
            validation.Append(new SpreadsheetLib.Formula2(options.Formula2));
        }

        validations.Append(validation);
        validations.Count = Convert.ToUInt32(validations.Count());
    }

    internal void AddConditionalFormatting(string reference, ConditionalFormattingOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);

        var conditionalFormatting = new SpreadsheetLib.ConditionalFormatting
        {
            SequenceOfReferences = new ListValue<StringValue> { InnerText = reference }
        };

        var rule = new SpreadsheetLib.ConditionalFormattingRule
        {
            Priority = Convert.ToInt32(GetNextConditionalFormattingPriority())
        };

        switch (options.Type)
        {
            case ConditionalFormattingType.GreaterThan:
                rule.Type = SpreadsheetLib.ConditionalFormatValues.CellIs;
                rule.Operator = SpreadsheetLib.ConditionalFormattingOperatorValues.GreaterThan;
                rule.FormatId = Spreadsheet.GetOrCreateDifferentialFormat(options.Style!);
                rule.Append(new SpreadsheetLib.Formula(options.Formula ?? throw new InvalidOperationException("Conditional formatting formula is required.")));
                break;
            case ConditionalFormattingType.LessThan:
                rule.Type = SpreadsheetLib.ConditionalFormatValues.CellIs;
                rule.Operator = SpreadsheetLib.ConditionalFormattingOperatorValues.LessThan;
                rule.FormatId = Spreadsheet.GetOrCreateDifferentialFormat(options.Style!);
                rule.Append(new SpreadsheetLib.Formula(options.Formula ?? throw new InvalidOperationException("Conditional formatting formula is required.")));
                break;
            case ConditionalFormattingType.ContainsText:
                rule.Type = SpreadsheetLib.ConditionalFormatValues.ContainsText;
                rule.Text = options.Text;
                rule.FormatId = Spreadsheet.GetOrCreateDifferentialFormat(options.Style!);
                rule.Append(new SpreadsheetLib.Formula($"NOT(ISERROR(SEARCH(\"{EscapeFormulaString(options.Text)}\",{reference.Split(':')[0]})))"));
                break;
            case ConditionalFormattingType.DuplicateValues:
                rule.Type = SpreadsheetLib.ConditionalFormatValues.DuplicateValues;
                rule.FormatId = Spreadsheet.GetOrCreateDifferentialFormat(options.Style!);
                break;
            case ConditionalFormattingType.TwoColorScale:
                rule.Type = SpreadsheetLib.ConditionalFormatValues.ColorScale;
                rule.Append(
                    new SpreadsheetLib.ColorScale(
                        new SpreadsheetLib.ConditionalFormatValueObject { Type = SpreadsheetLib.ConditionalFormatValueObjectValues.Min },
                        new SpreadsheetLib.ConditionalFormatValueObject { Type = SpreadsheetLib.ConditionalFormatValueObjectValues.Max },
                        new SpreadsheetLib.Color { Rgb = Utils.ArgbHexConverter(options.MinimumColor ?? Color.LightGreen) },
                        new SpreadsheetLib.Color { Rgb = Utils.ArgbHexConverter(options.MaximumColor ?? Color.DarkGreen) }
                    )
                );
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(options));
        }

        conditionalFormatting.Append(rule);
        InsertAfterMergeCells(conditionalFormatting);
    }

    internal void SetCellHyperlink(Cell cell, string target, string? displayText)
    {
        if (string.IsNullOrWhiteSpace(target))
        {
            throw new ArgumentException("Hyperlink target cannot be null or empty.", nameof(target));
        }

        if (!string.IsNullOrEmpty(displayText))
        {
            cell.SetValue(displayText);
        }

        var worksheetHyperlinks = WorksheetElement.GetFirstChild<SpreadsheetLib.Hyperlinks>();
        if (worksheetHyperlinks == null)
        {
            worksheetHyperlinks = new SpreadsheetLib.Hyperlinks();
            InsertAfterDataValidations(worksheetHyperlinks);
        }

        var existingHyperlink = worksheetHyperlinks.Elements<SpreadsheetLib.Hyperlink>()
            .FirstOrDefault(hyperlink => hyperlink.Reference?.Value == cell.CellReference);

        if (existingHyperlink?.Id?.Value is { Length: > 0 } existingRelationshipId)
        {
            WorksheetPart.DeleteReferenceRelationship(existingRelationshipId);
        }

        existingHyperlink?.Remove();

        SpreadsheetLib.Hyperlink hyperlink;
        if (Uri.TryCreate(target, UriKind.Absolute, out var absoluteUri))
        {
            var relationship = WorksheetPart.AddHyperlinkRelationship(absoluteUri, true);
            hyperlink = new SpreadsheetLib.Hyperlink
            {
                Reference = cell.CellReference,
                Id = relationship.Id
            };
        }
        else
        {
            hyperlink = new SpreadsheetLib.Hyperlink
            {
                Reference = cell.CellReference,
                Location = target.TrimStart('#')
            };
        }

        worksheetHyperlinks.Append(hyperlink);
    }

    internal string? GetCellHyperlink(string cellReference)
    {
        var hyperlink = WorksheetElement.GetFirstChild<SpreadsheetLib.Hyperlinks>()?
            .Elements<SpreadsheetLib.Hyperlink>()
            .FirstOrDefault(current => current.Reference?.Value == cellReference);

        if (hyperlink == null)
        {
            return null;
        }

        if (hyperlink.Id?.Value is { Length: > 0 } relationshipId)
        {
            return WorksheetPart.HyperlinkRelationships.FirstOrDefault(relationship => relationship.Id == relationshipId)?.Uri.ToString();
        }

        return hyperlink.Location?.Value;
    }

    internal void SetCellComment(Cell cell, string text, string? author)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            throw new ArgumentException("Comment text cannot be null or empty.", nameof(text));
        }

        var commentsPart = WorksheetPart.WorksheetCommentsPart ?? WorksheetPart.AddNewPart<WorksheetCommentsPart>();
        var comments = commentsPart.Comments ??= new SpreadsheetLib.Comments(new SpreadsheetLib.Authors(), new SpreadsheetLib.CommentList());
        var authors = comments.Authors ?? comments.AppendChild(new SpreadsheetLib.Authors());
        var commentList = comments.CommentList ?? comments.AppendChild(new SpreadsheetLib.CommentList());

        author ??= "OfficeDocuments";
        var authorIndex = authors.Elements<SpreadsheetLib.Author>()
            .Select((item, index) => new { item, index })
            .FirstOrDefault(item => string.Equals(item.item.Text, author, StringComparison.Ordinal))?.index;

        if (authorIndex == null)
        {
            authors.Append(new SpreadsheetLib.Author(author));
            authorIndex = authors.Count() - 1;
        }

        var comment = commentList.Elements<SpreadsheetLib.Comment>()
            .FirstOrDefault(item => item.Reference?.Value == cell.CellReference);

        if (comment == null)
        {
            comment = new SpreadsheetLib.Comment
            {
                Reference = cell.CellReference,
                AuthorId = Convert.ToUInt32(authorIndex.Value)
            };
            commentList.Append(comment);
        }
        else
        {
            comment.AuthorId = Convert.ToUInt32(authorIndex.Value);
            comment.RemoveAllChildren<SpreadsheetLib.CommentText>();
        }

        comment.Append(
            new SpreadsheetLib.CommentText(
                new SpreadsheetLib.Run(
                    new SpreadsheetLib.RunProperties(),
                    new SpreadsheetLib.Text(text) { Space = SpaceProcessingModeValues.Preserve }
                )
            )
        );

        comments.Save();
        UpdateCommentVml(commentList.Elements<SpreadsheetLib.Comment>().Select(item => item.Reference?.Value).OfType<string>());
    }

    internal string? GetCellComment(string cellReference)
    {
        return WorksheetPart.WorksheetCommentsPart?.Comments?
            .CommentList?
            .Elements<SpreadsheetLib.Comment>()
            .FirstOrDefault(comment => comment.Reference?.Value == cellReference)?
            .CommentText?
            .InnerText;
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

    private uint GetNextConditionalFormattingPriority()
    {
        var priorities = WorksheetElement.Elements<SpreadsheetLib.ConditionalFormatting>()
            .SelectMany(item => item.Elements<SpreadsheetLib.ConditionalFormattingRule>())
            .Select(item => item.Priority?.Value ?? 0);

        return Convert.ToUInt32(priorities.DefaultIfEmpty().Max() + 1);
    }

    private void InsertAfterMergeCells(OpenXmlElement element)
    {
        var lastConditionalFormatting = WorksheetElement.Elements<SpreadsheetLib.ConditionalFormatting>().LastOrDefault();
        if (lastConditionalFormatting != null)
        {
            WorksheetElement.InsertAfter(element, lastConditionalFormatting);
            return;
        }

        var mergeCells = WorksheetElement.GetFirstChild<SpreadsheetLib.MergeCells>();
        if (mergeCells != null)
        {
            WorksheetElement.InsertAfter(element, mergeCells);
            return;
        }

        var autoFilter = WorksheetElement.GetFirstChild<SpreadsheetLib.AutoFilter>();
        if (autoFilter != null)
        {
            WorksheetElement.InsertAfter(element, autoFilter);
            return;
        }

        WorksheetElement.InsertAfter(element, Element);
    }

    private void InsertAfterConditionalFormatting(OpenXmlElement element)
    {
        var lastConditionalFormatting = WorksheetElement.Elements<SpreadsheetLib.ConditionalFormatting>().LastOrDefault();
        if (lastConditionalFormatting != null)
        {
            WorksheetElement.InsertAfter(element, lastConditionalFormatting);
            return;
        }

        var mergeCells = WorksheetElement.GetFirstChild<SpreadsheetLib.MergeCells>();
        if (mergeCells != null)
        {
            WorksheetElement.InsertAfter(element, mergeCells);
            return;
        }

        var autoFilter = WorksheetElement.GetFirstChild<SpreadsheetLib.AutoFilter>();
        if (autoFilter != null)
        {
            WorksheetElement.InsertAfter(element, autoFilter);
            return;
        }

        WorksheetElement.InsertAfter(element, Element);
    }

    private void InsertAfterDataValidations(OpenXmlElement element)
    {
        var dataValidations = WorksheetElement.GetFirstChild<SpreadsheetLib.DataValidations>();
        if (dataValidations != null)
        {
            WorksheetElement.InsertAfter(element, dataValidations);
            return;
        }

        var lastConditionalFormatting = WorksheetElement.Elements<SpreadsheetLib.ConditionalFormatting>().LastOrDefault();
        if (lastConditionalFormatting != null)
        {
            WorksheetElement.InsertAfter(element, lastConditionalFormatting);
            return;
        }

        var mergeCells = WorksheetElement.GetFirstChild<SpreadsheetLib.MergeCells>();
        if (mergeCells != null)
        {
            WorksheetElement.InsertAfter(element, mergeCells);
            return;
        }

        var autoFilter = WorksheetElement.GetFirstChild<SpreadsheetLib.AutoFilter>();
        if (autoFilter != null)
        {
            WorksheetElement.InsertAfter(element, autoFilter);
            return;
        }

        WorksheetElement.InsertAfter(element, Element);
    }

    private void UpdateCommentVml(IEnumerable<string> references)
    {
        var vmlPart = WorksheetPart.VmlDrawingParts.FirstOrDefault() ?? WorksheetPart.AddNewPart<VmlDrawingPart>();
        var vmlRelationshipId = WorksheetPart.GetIdOfPart(vmlPart);

        var legacyDrawing = WorksheetElement.GetFirstChild<SpreadsheetLib.LegacyDrawing>();
        if (legacyDrawing == null)
        {
            WorksheetElement.Append(new SpreadsheetLib.LegacyDrawing { Id = vmlRelationshipId });
        }
        else
        {
            legacyDrawing.Id = vmlRelationshipId;
        }

        using var stream = vmlPart.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, Encoding.UTF8);
        writer.Write(BuildCommentVml(references));
    }

    private static string BuildCommentVml(IEnumerable<string> references)
    {
        var builder = new StringBuilder();
        builder.AppendLine("""<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">""");
        builder.AppendLine("""<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>""");
        builder.AppendLine("""<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>""");

        var shapeId = 1025;
        foreach (var reference in references)
        {
            var (rowIndex, columnIndex) = reference.GetExcelCellIndex();
            var zeroBasedRow = rowIndex - 1;
            var zeroBasedColumn = columnIndex - 1;
            var anchor = $"{zeroBasedColumn}, 15, {zeroBasedRow}, 2, {zeroBasedColumn + 3}, 15, {zeroBasedRow + 4}, 4";

            builder.AppendLine(
                $"""<v:shape id="_x0000_s{shapeId++}" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:5pt;width:104pt;height:64pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>{anchor}</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>{zeroBasedRow}</x:Row><x:Column>{zeroBasedColumn}</x:Column></x:ClientData></v:shape>"""
            );
        }

        builder.AppendLine("</xml>");
        return builder.ToString();
    }

    private static string EscapeFormulaString(string? value)
    {
        return (value ?? string.Empty).Replace("\"", "\"\"");
    }

    public void AddImage(string filePath, uint fromColumn, uint fromRow, uint toColumn, uint toRow)
    {
        ArgumentException.ThrowIfNullOrEmpty(filePath);
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException("Image file not found.", filePath);
        }

        var imageType = DetectImageType(filePath);
        using var stream = File.OpenRead(filePath);
        AddImage(stream, imageType, fromColumn, fromRow, toColumn, toRow);
    }

    public void AddImage(Stream imageStream, ImageType imageType, uint fromColumn, uint fromRow, uint toColumn, uint toRow)
    {
        ArgumentNullException.ThrowIfNull(imageStream);
        if (fromColumn < 1)
        {
            throw new ArgumentException("fromColumn must be at least 1.", nameof(fromColumn));
        }

        if (fromRow < 1)
        {
            throw new ArgumentException("fromRow must be at least 1.", nameof(fromRow));
        }

        if (toColumn < fromColumn)
        {
            throw new ArgumentException("toColumn must be greater than or equal to fromColumn.", nameof(toColumn));
        }

        if (toRow < fromRow)
        {
            throw new ArgumentException("toRow must be greater than or equal to fromRow.", nameof(toRow));
        }

        var drawingsPart = WorksheetPart.DrawingsPart ?? WorksheetPart.AddNewPart<DrawingsPart>();
        var imagePart = drawingsPart.AddImagePart(ToImagePartType(imageType));
        imagePart.FeedData(imageStream);
        var imageRelId = drawingsPart.GetIdOfPart(imagePart);

        drawingsPart.WorksheetDrawing ??= new XdrSpr.WorksheetDrawing();
        var existingCount = drawingsPart.WorksheetDrawing.Elements<XdrSpr.TwoCellAnchor>().Count()
            + drawingsPart.WorksheetDrawing.Elements<XdrSpr.OneCellAnchor>().Count();
        var pictureId = (uint)(existingCount + 1);

        drawingsPart.WorksheetDrawing.Append(BuildTwoCellAnchor(imageRelId, pictureId, fromColumn, fromRow, toColumn, toRow));
        EnsureDrawingElement(WorksheetPart.GetIdOfPart(drawingsPart));
    }

    private static XdrSpr.TwoCellAnchor BuildTwoCellAnchor(
        string imageRelId, uint pictureId, uint fromColumn, uint fromRow, uint toColumn, uint toRow)
    {
        var anchor = new XdrSpr.TwoCellAnchor();
        anchor.Append(new XdrSpr.FromMarker(
            new XdrSpr.ColumnId((fromColumn - 1).ToString()),
            new XdrSpr.ColumnOffset("0"),
            new XdrSpr.RowId((fromRow - 1).ToString()),
            new XdrSpr.RowOffset("0")
        ));
        anchor.Append(new XdrSpr.ToMarker(
            new XdrSpr.ColumnId(toColumn.ToString()),
            new XdrSpr.ColumnOffset("0"),
            new XdrSpr.RowId(toRow.ToString()),
            new XdrSpr.RowOffset("0")
        ));
        var picture = new XdrSpr.Picture();
        picture.Append(new XdrSpr.NonVisualPictureProperties(
            new XdrSpr.NonVisualDrawingProperties { Id = pictureId, Name = $"Image{pictureId}" },
            new XdrSpr.NonVisualPictureDrawingProperties()
        ));
        picture.Append(new XdrSpr.BlipFill(
            new Drw.Blip { Embed = imageRelId },
            new Drw.Stretch(new Drw.FillRectangle())
        ));
        picture.Append(new XdrSpr.ShapeProperties(
            new Drw.Transform2D(
                new Drw.Offset { X = 0, Y = 0 },
                new Drw.Extents { Cx = 0, Cy = 0 }
            ),
            new Drw.PresetGeometry(new Drw.AdjustValueList()) { Preset = Drw.ShapeTypeValues.Rectangle }
        ));
        anchor.Append(picture);
        anchor.Append(new XdrSpr.ClientData());
        return anchor;
    }

    private void EnsureDrawingElement(string drawingRelId)
    {
        if (WorksheetElement.GetFirstChild<SpreadsheetLib.Drawing>() != null)
        {
            return;
        }

        var drawing = new SpreadsheetLib.Drawing { Id = drawingRelId };
        var legacyDrawing = WorksheetElement.GetFirstChild<SpreadsheetLib.LegacyDrawing>();
        if (legacyDrawing != null)
        {
            WorksheetElement.InsertBefore(drawing, legacyDrawing);
            return;
        }

        var tableParts = WorksheetElement.GetFirstChild<SpreadsheetLib.TableParts>();
        if (tableParts != null)
        {
            WorksheetElement.InsertBefore(drawing, tableParts);
            return;
        }

        WorksheetElement.AppendChild(drawing);
    }

    private static PartTypeInfo ToImagePartType(ImageType imageType) => imageType switch
    {
        ImageType.Png => ImagePartType.Png,
        ImageType.Jpeg => ImagePartType.Jpeg,
        ImageType.Gif => ImagePartType.Gif,
        ImageType.Bmp => ImagePartType.Bmp,
        ImageType.Tiff => ImagePartType.Tiff,
        _ => throw new ArgumentOutOfRangeException(nameof(imageType))
    };

    private static ImageType DetectImageType(string filePath)
    {
        return Path.GetExtension(filePath).ToLowerInvariant() switch
        {
            ".png" => ImageType.Png,
            ".jpg" or ".jpeg" => ImageType.Jpeg,
            ".gif" => ImageType.Gif,
            ".bmp" => ImageType.Bmp,
            ".tiff" or ".tif" => ImageType.Tiff,
            var ext => throw new ArgumentException($"Unsupported image format: '{ext}'.", nameof(filePath))
        };
    }
}
