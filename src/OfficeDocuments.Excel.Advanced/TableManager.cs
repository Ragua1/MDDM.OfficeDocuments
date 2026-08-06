using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Excel.DataClasses;
using OfficeDocuments.Excel.Extensions;
using OfficeDocuments.Excel.Interfaces;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.Advanced;

/// <summary>
/// Owns structured-table create/lookup/lifecycle over the workbook. Depends only on the
/// workbook part and worksheet-lookup seams so it can move to an advanced layer later.
/// </summary>
internal sealed class TableManager(
    WorkbookPart workbookPart,
    Func<string, Worksheet> getWorksheetOrThrow,
    Func<IEnumerable<Worksheet>> getAllWorksheets)
{
    public ITableInfo AddTable(string worksheetName, ICell startCell, ICell endCell, List<string> columnsName)
    {
        return AddTableCore(worksheetName, startCell, endCell, columnsName, options: null);
    }

    public ITableInfo AddTable(IRange range, List<string> columnsName, TableCreateOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(range);

        var startCell = range.Worksheet.AddCellOnIndex(range.FromColumn, range.FromRow);
        var endCell = range.Worksheet.AddCellOnIndex(range.ToColumn, range.ToRow);
        return AddTableCore(range.Worksheet.Name, startCell, endCell, columnsName, options);
    }

    public ITableInfo? GetTable(string worksheetName, string tableName)
    {
        ArgumentException.ThrowIfNullOrEmpty(worksheetName);
        ArgumentException.ThrowIfNullOrEmpty(tableName);

        var worksheet = getWorksheetOrThrow(worksheetName);
        var tablePart = FindTableDefinitionPart(worksheet.WorksheetPart, tableName);
        return tablePart == null ? null : ToTableInfo(worksheetName, tablePart);
    }

    public IEnumerable<ITableInfo> GetTables(string worksheetName)
    {
        ArgumentException.ThrowIfNullOrEmpty(worksheetName);

        var worksheet = getWorksheetOrThrow(worksheetName);
        return worksheet.WorksheetPart.TableDefinitionParts
            .Select(part => ToTableInfo(worksheetName, part))
            .ToArray();
    }

    public IEnumerable<ITableInfo> GetTables()
    {
        return getAllWorksheets()
            .SelectMany(ws => ws.WorksheetPart.TableDefinitionParts
                .Select(part => ToTableInfo(ws.Name, part)))
            .ToArray();
    }

    public void RenameTable(string worksheetName, string tableName, string newName)
    {
        ArgumentException.ThrowIfNullOrEmpty(worksheetName);
        ArgumentException.ThrowIfNullOrEmpty(tableName);
        ArgumentException.ThrowIfNullOrEmpty(newName);

        var allTables = GetTables();
        if (allTables.Any(t => string.Equals(t.Name, newName, StringComparison.OrdinalIgnoreCase) && !string.Equals(t.Name, tableName, StringComparison.OrdinalIgnoreCase)))
        {
            throw new ArgumentException($"A table named '{newName}' already exists in the workbook.", nameof(newName));
        }

        var worksheet = getWorksheetOrThrow(worksheetName);
        var tablePart = FindTableDefinitionPart(worksheet.WorksheetPart, tableName)
            ?? throw new ArgumentException($"Table '{tableName}' not found on worksheet '{worksheetName}'.", nameof(tableName));
        var table = tablePart.Table ?? throw new InvalidOperationException($"Table '{tableName}' does not contain a table definition.");

        table.Name = newName;
        table.DisplayName = newName;
    }

    public void ResizeTable(string worksheetName, string tableName, IRange newRange)
    {
        ArgumentException.ThrowIfNullOrEmpty(worksheetName);
        ArgumentException.ThrowIfNullOrEmpty(tableName);
        ArgumentNullException.ThrowIfNull(newRange);

        var worksheet = getWorksheetOrThrow(worksheetName);
        var tablePart = FindTableDefinitionPart(worksheet.WorksheetPart, tableName)
            ?? throw new ArgumentException($"Table '{tableName}' not found on worksheet '{worksheetName}'.", nameof(tableName));

        var table = tablePart.Table ?? throw new InvalidOperationException($"Table '{tableName}' does not contain a table definition.");
        var existingColumnCount = (int)(table.TableColumns?.Count?.Value ?? 0);
        var newColumnCount = (int)(newRange.ToColumn - newRange.FromColumn + 1);
        if (newColumnCount != existingColumnCount)
        {
            throw new ArgumentException($"Cannot resize table '{tableName}': new range has {newColumnCount} columns but table has {existingColumnCount} columns.", nameof(newRange));
        }

        var startRef = CellExtension.GetExcelCellReference(newRange.FromColumn, newRange.FromRow);
        var endRef = CellExtension.GetExcelCellReference(newRange.ToColumn, newRange.ToRow);
        var newRef = $"{startRef}:{endRef}";
        table.Reference = newRef;
        if (table.AutoFilter != null)
        {
            table.AutoFilter.Reference = newRef;
        }
    }

    public void RemoveTable(string worksheetName, string tableName)
    {
        ArgumentException.ThrowIfNullOrEmpty(worksheetName);
        ArgumentException.ThrowIfNullOrEmpty(tableName);

        var worksheet = getWorksheetOrThrow(worksheetName);
        var tablePart = FindTableDefinitionPart(worksheet.WorksheetPart, tableName)
            ?? throw new ArgumentException($"Table '{tableName}' not found on worksheet '{worksheetName}'.", nameof(tableName));

        var tableRelId = worksheet.WorksheetPart.GetIdOfPart(tablePart);
        var tableParts = worksheet.WorksheetElement.GetFirstChild<SpreadsheetLib.TableParts>();
        if (tableParts != null)
        {
            var tablePartRef = tableParts.Elements<SpreadsheetLib.TablePart>()
                .FirstOrDefault(tp => tp.Id?.Value == tableRelId);
            tablePartRef?.Remove();

            if (!tableParts.Elements<SpreadsheetLib.TablePart>().Any())
            {
                tableParts.Remove();
            }
            else
            {
                tableParts.Count = Convert.ToUInt32(tableParts.Elements<SpreadsheetLib.TablePart>().Count());
            }
        }

        worksheet.WorksheetPart.DeletePart(tablePart);
    }

    private ITableInfo AddTableCore(string worksheetName, ICell startCell, ICell endCell, List<string> columnsName, TableCreateOptions? options)
    {
        ArgumentException.ThrowIfNullOrEmpty(worksheetName);
        ArgumentNullException.ThrowIfNull(startCell);
        ArgumentNullException.ThrowIfNull(endCell);
        ArgumentNullException.ThrowIfNull(columnsName);

        if (columnsName.Count == 0)
        {
            throw new ArgumentException("Column names list cannot be empty.", nameof(columnsName));
        }

        if (columnsName.Any(string.IsNullOrWhiteSpace))
        {
            throw new ArgumentException("Table column names cannot be null or empty.", nameof(columnsName));
        }

        if (startCell.RowIndex > endCell.RowIndex || startCell.ColumnIndex > endCell.ColumnIndex)
        {
            throw new ArgumentException("Invalid table definition: start cell must be before end cell.");
        }

        var expectedColumnCount = endCell.ColumnIndex - startCell.ColumnIndex + 1;
        if (columnsName.Count != expectedColumnCount)
        {
            throw new ArgumentException("The number of table columns must match the table width.", nameof(columnsName));
        }

        var worksheet = getWorksheetOrThrow(worksheetName);
        var worksheetPart = worksheet.WorksheetPart;
        var tableIndex = workbookPart.WorksheetParts.SelectMany(part => part.TableDefinitionParts).Count() + 1;
        var autoName = $"Table{tableIndex}";
        var tableName = options?.TableName ?? autoName;
        var displayName = options?.DisplayName ?? tableName;

        var styleOptions = options?.Style;
        var tableRef = $"{startCell.CellReference}:{endCell.CellReference}";
        var table = new SpreadsheetLib.Table
        {
            Id = (uint)tableIndex,
            Name = tableName,
            DisplayName = displayName,
            Reference = tableRef,
            TotalsRowShown = false,
            TableColumns = new SpreadsheetLib.TableColumns { Count = Convert.ToUInt32(columnsName.Count) },
            AutoFilter = new SpreadsheetLib.AutoFilter { Reference = tableRef },
            TableStyleInfo = new SpreadsheetLib.TableStyleInfo
            {
                Name = styleOptions?.StyleName ?? "TableStyleMedium2",
                ShowFirstColumn = styleOptions?.ShowFirstColumn ?? false,
                ShowLastColumn = styleOptions?.ShowLastColumn ?? false,
                ShowRowStripes = styleOptions?.ShowBandedRows ?? true,
                ShowColumnStripes = styleOptions?.ShowBandedColumns ?? false
            }
        };

        for (var index = 0; index < columnsName.Count; index++)
        {
            table.TableColumns.Append(new SpreadsheetLib.TableColumn
            {
                Id = (uint)index + 1,
                Name = columnsName[index]
            });
        }

        var tablePart = worksheetPart.AddNewPart<TableDefinitionPart>();
        tablePart.Table = table;
        var tableRelationshipId = worksheetPart.GetIdOfPart(tablePart);

        var tableParts = worksheet.WorksheetElement.GetFirstChild<SpreadsheetLib.TableParts>();
        if (tableParts == null)
        {
            tableParts = worksheet.WorksheetElement.AppendChild(new SpreadsheetLib.TableParts());
        }

        tableParts.Append(new SpreadsheetLib.TablePart { Id = tableRelationshipId });
        tableParts.Count = Convert.ToUInt32(tableParts.Elements<SpreadsheetLib.TablePart>().Count());

        return new TableInfo(tableName, displayName, tableRef, columnsName.AsReadOnly(), worksheetName);
    }

    private static TableDefinitionPart? FindTableDefinitionPart(WorksheetPart worksheetPart, string tableName)
    {
        return worksheetPart.TableDefinitionParts
            .FirstOrDefault(part => string.Equals(part.Table?.Name, tableName, StringComparison.OrdinalIgnoreCase));
    }

    private static ITableInfo ToTableInfo(string worksheetName, TableDefinitionPart part)
    {
        var table = part.Table ?? throw new InvalidOperationException("The table definition part does not contain a table.");
        var columnNames = table.TableColumns?
            .Elements<SpreadsheetLib.TableColumn>()
            .OrderBy(col => col.Id?.Value)
            .Select(col => col.Name?.Value ?? string.Empty)
            .ToList() ?? [];

        return new TableInfo(
            table.Name?.Value ?? string.Empty,
            table.DisplayName?.Value ?? string.Empty,
            table.Reference?.Value ?? string.Empty,
            columnNames.AsReadOnly(),
            worksheetName);
    }
}
