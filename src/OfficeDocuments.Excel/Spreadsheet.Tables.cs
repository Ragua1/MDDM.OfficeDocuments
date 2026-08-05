using OfficeDocuments.Excel.DataClasses;
using OfficeDocuments.Excel.Interfaces;
using OfficeDocuments.Excel.Options;

namespace OfficeDocuments.Excel;

public partial class Spreadsheet
{
    private TableManager? _tableManager;

    private TableManager TableManager =>
        _tableManager ??= new TableManager(WorkbookPartInternal, GetWorksheetOrThrow, () => _worksheets.OfType<Worksheet>());

    public ITableInfo AddTable(string worksheetName, ICell startCell, ICell endCell, List<string> columnsName) =>
        TableManager.AddTable(worksheetName, startCell, endCell, columnsName);

    public ITableInfo AddTable(IRange range, List<string> columnsName, TableCreateOptions? options = null) =>
        TableManager.AddTable(range, columnsName, options);

    public ITableInfo? GetTable(string worksheetName, string tableName) =>
        TableManager.GetTable(worksheetName, tableName);

    public IEnumerable<ITableInfo> GetTables(string worksheetName) =>
        TableManager.GetTables(worksheetName);

    public IEnumerable<ITableInfo> GetTables() =>
        TableManager.GetTables();

    public void RenameTable(string worksheetName, string tableName, string newName) =>
        TableManager.RenameTable(worksheetName, tableName, newName);

    public void ResizeTable(string worksheetName, string tableName, IRange newRange) =>
        TableManager.ResizeTable(worksheetName, tableName, newRange);

    public void RemoveTable(string worksheetName, string tableName) =>
        TableManager.RemoveTable(worksheetName, tableName);
}
