namespace OfficeDocuments.Excel.Advanced;

internal sealed class TableInfo : ITableInfo
{
    public string Name { get; }
    public string DisplayName { get; }
    public string Reference { get; }
    public int ColumnCount { get; }
    public IReadOnlyList<string> ColumnNames { get; }
    public string WorksheetName { get; }

    internal TableInfo(string name, string displayName, string reference, IReadOnlyList<string> columnNames, string worksheetName)
    {
        Name = name;
        DisplayName = displayName;
        Reference = reference;
        ColumnNames = columnNames;
        ColumnCount = columnNames.Count;
        WorksheetName = worksheetName;
    }
}
