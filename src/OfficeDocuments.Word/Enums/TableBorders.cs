namespace OfficeDocuments.Word.Enums;

/// <summary>
/// Which edges of a table get borders.
/// </summary>
public enum TableBorders
{
    /// <summary>
    /// No borders at all, overriding any the table style would apply.
    /// </summary>
    None,
    /// <summary>
    /// The outer edges only.
    /// </summary>
    Outline,
    /// <summary>
    /// Outer edges plus the lines between rows and columns.
    /// </summary>
    All,
}
