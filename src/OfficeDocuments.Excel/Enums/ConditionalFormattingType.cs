namespace OfficeDocuments.Excel.Enums;

/// <summary>
/// The conditional-formatting rule kinds this library can write.
/// </summary>
/// <remarks>
/// Build a rule through the factory methods on <see cref="Options.ConditionalFormattingOptions"/>
/// rather than setting the type by hand — each kind needs a different combination of the remaining
/// options.
/// </remarks>
public enum ConditionalFormattingType
{
    /// <summary>
    /// Formats cells whose value is greater than a threshold.
    /// </summary>
    GreaterThan = 0,

    /// <summary>
    /// Formats cells whose value is less than a threshold.
    /// </summary>
    LessThan = 1,

    /// <summary>
    /// Formats cells whose text contains a substring.
    /// </summary>
    ContainsText = 2,

    /// <summary>
    /// Formats cells whose value appears more than once in the range.
    /// </summary>
    DuplicateValues = 3,

    /// <summary>
    /// Shades the range along a gradient between a minimum and a maximum colour. This is the one
    /// kind that takes no style, because the colours are the formatting.
    /// </summary>
    TwoColorScale = 4
}
