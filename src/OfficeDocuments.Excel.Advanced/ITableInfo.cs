namespace OfficeDocuments.Excel.Advanced;

/// <summary>
/// Read-only metadata about a structured table in a worksheet.
/// </summary>
public interface ITableInfo
{
    /// <summary>Gets the internal table name used in XML references.</summary>
    string Name { get; }

    /// <summary>Gets the display name shown in the Excel UI.</summary>
    string DisplayName { get; }

    /// <summary>Gets the table reference, e.g. "A1:C5".</summary>
    string Reference { get; }

    /// <summary>Gets the number of columns in the table.</summary>
    int ColumnCount { get; }

    /// <summary>Gets the ordered column names of the table.</summary>
    IReadOnlyList<string> ColumnNames { get; }

    /// <summary>Gets the name of the worksheet that contains the table.</summary>
    string WorksheetName { get; }
}
