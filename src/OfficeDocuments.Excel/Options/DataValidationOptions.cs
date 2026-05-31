using System.Collections.Generic;
using OfficeDocuments.Excel.Enums;

namespace OfficeDocuments.Excel.Options;

public sealed record DataValidationOptions
{
    public required DataValidationType Type { get; init; }
    public DataValidationOperator? Operator { get; init; }
    public required string Formula1 { get; init; }
    public string? Formula2 { get; init; }
    public bool AllowBlank { get; init; } = true;
    public bool ShowDropDown { get; init; } = true;
    public string? PromptTitle { get; init; }
    public string? Prompt { get; init; }
    public string? ErrorTitle { get; init; }
    public string? Error { get; init; }

    public static DataValidationOptions List(IEnumerable<string> values)
    {
        ArgumentNullException.ThrowIfNull(values);

        var normalizedValues = values
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value.Replace("\"", "\"\""))
            .ToArray();

        if (normalizedValues.Length == 0)
        {
            throw new ArgumentException("Validation list must contain at least one value.", nameof(values));
        }

        return new DataValidationOptions
        {
            Type = DataValidationType.List,
            Formula1 = $"\"{string.Join(",", normalizedValues)}\""
        };
    }

    public static DataValidationOptions WholeNumber(DataValidationOperator @operator, string formula1, string? formula2 = null)
        => CreateComparison(DataValidationType.Whole, @operator, formula1, formula2);

    public static DataValidationOptions DecimalNumber(DataValidationOperator @operator, string formula1, string? formula2 = null)
        => CreateComparison(DataValidationType.Decimal, @operator, formula1, formula2);

    public static DataValidationOptions Date(DataValidationOperator @operator, string formula1, string? formula2 = null)
        => CreateComparison(DataValidationType.Date, @operator, formula1, formula2);

    public static DataValidationOptions Custom(string formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
        {
            throw new ArgumentException("Validation formula cannot be null or empty.", nameof(formula));
        }

        return new DataValidationOptions
        {
            Type = DataValidationType.Custom,
            Formula1 = formula
        };
    }

    private static DataValidationOptions CreateComparison(DataValidationType type, DataValidationOperator @operator, string formula1, string? formula2)
    {
        if (string.IsNullOrWhiteSpace(formula1))
        {
            throw new ArgumentException("Validation formula cannot be null or empty.", nameof(formula1));
        }

        if ((@operator == DataValidationOperator.Between || @operator == DataValidationOperator.NotBetween)
            && string.IsNullOrWhiteSpace(formula2))
        {
            throw new ArgumentException("Validation formula2 is required for between operators.", nameof(formula2));
        }

        return new DataValidationOptions
        {
            Type = type,
            Operator = @operator,
            Formula1 = formula1,
            Formula2 = formula2
        };
    }
}
