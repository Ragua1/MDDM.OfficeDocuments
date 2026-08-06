namespace OfficeDocuments.Excel.Advanced;

/// <summary>
/// Options for creating a structured table.
/// </summary>
public sealed class TableCreateOptions
{
    /// <summary>
    /// Gets the internal table name. When null, a unique name is generated automatically.
    /// </summary>
    public string? TableName { get; init; }

    /// <summary>
    /// Gets the display name shown in the Excel UI. Defaults to <see cref="TableName"/> when null.
    /// </summary>
    public string? DisplayName { get; init; }

    /// <summary>
    /// Gets the style and behavior options for the table.
    /// Uses library defaults when null.
    /// </summary>
    public TableStyleOptions? Style { get; init; }
}
