namespace OfficeDocuments.Excel.Enums;

/// <summary>
/// What kind of entry a data-validation rule accepts.
/// </summary>
/// <remarks>
/// Build a rule through the factory methods on <see cref="Options.DataValidationOptions"/>; each
/// one pairs the type with the bounds and operator that type actually uses.
/// </remarks>
public enum DataValidationType
{
    /// <summary>
    /// Accepts only values from a fixed list, shown as a drop-down in Excel.
    /// </summary>
    List = 0,

    /// <summary>
    /// Accepts whole numbers within the bounds set by the rule's operator.
    /// </summary>
    Whole = 1,

    /// <summary>
    /// Accepts decimal numbers within the bounds set by the rule's operator.
    /// </summary>
    Decimal = 2,

    /// <summary>
    /// Accepts dates within the bounds set by the rule's operator.
    /// </summary>
    Date = 3,

    /// <summary>
    /// Accepts whatever an Excel formula evaluates to <see langword="true"/> for. The operator is
    /// not used.
    /// </summary>
    Custom = 4
}
