namespace OfficeDocuments.Excel.Options;

/// <summary>
/// Style and behavior options for a structured table.
/// </summary>
public sealed class TableStyleOptions
{
    /// <summary>
    /// Gets the table style preset name. Defaults to "TableStyleMedium2".
    /// </summary>
    public string StyleName { get; init; } = "TableStyleMedium2";

    /// <summary>
    /// Gets whether to highlight the first column. Defaults to false.
    /// </summary>
    public bool ShowFirstColumn { get; init; } = false;

    /// <summary>
    /// Gets whether to highlight the last column. Defaults to false.
    /// </summary>
    public bool ShowLastColumn { get; init; } = false;

    /// <summary>
    /// Gets whether to display banded rows. Defaults to true.
    /// </summary>
    public bool ShowBandedRows { get; init; } = true;

    /// <summary>
    /// Gets whether to display banded columns. Defaults to false.
    /// </summary>
    public bool ShowBandedColumns { get; init; } = false;
}
