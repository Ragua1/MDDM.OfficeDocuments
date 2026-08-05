namespace OfficeDocuments.Excel.Enums;

/// <summary>
/// The direction <see cref="Interfaces.IRange.SortByColumn"/> orders rows in.
/// </summary>
public enum SortDirection
{
    /// <summary>
    /// Smallest first: numbers ascending, text A to Z, earliest dates first.
    /// </summary>
    Ascending = 0,

    /// <summary>
    /// Largest first: numbers descending, text Z to A, latest dates first.
    /// </summary>
    Descending = 1
}
