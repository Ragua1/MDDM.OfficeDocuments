namespace OfficeDocuments.Excel.Enums;

/// <summary>
/// Compares a cell's entry against the bounds of a data-validation rule.
/// </summary>
/// <remarks>
/// <see cref="Between"/> and <see cref="NotBetween"/> are the only members that use a second bound;
/// for the rest the second value is ignored. The operator has no effect on a
/// <see cref="DataValidationType.List"/> or <see cref="DataValidationType.Custom"/> rule.
/// </remarks>
public enum DataValidationOperator
{
    /// <summary>
    /// Accepts values from the first bound to the second, inclusive.
    /// </summary>
    Between = 0,

    /// <summary>
    /// Accepts values outside the range from the first bound to the second.
    /// </summary>
    NotBetween = 1,

    /// <summary>
    /// Accepts only the first bound.
    /// </summary>
    Equal = 2,

    /// <summary>
    /// Accepts anything other than the first bound.
    /// </summary>
    NotEqual = 3,

    /// <summary>
    /// Accepts values strictly above the first bound.
    /// </summary>
    GreaterThan = 4,

    /// <summary>
    /// Accepts values strictly below the first bound.
    /// </summary>
    LessThan = 5,

    /// <summary>
    /// Accepts values at or above the first bound.
    /// </summary>
    GreaterThanOrEqual = 6,

    /// <summary>
    /// Accepts values at or below the first bound.
    /// </summary>
    LessThanOrEqual = 7
}
