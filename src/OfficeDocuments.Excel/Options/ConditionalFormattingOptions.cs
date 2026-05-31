using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Interfaces;
using Color = System.Drawing.Color;

namespace OfficeDocuments.Excel.Options;

public sealed record ConditionalFormattingOptions
{
    public required ConditionalFormattingType Type { get; init; }
    public string? Formula { get; init; }
    public string? Text { get; init; }
    public IStyle? Style { get; init; }
    public Color? MinimumColor { get; init; }
    public Color? MaximumColor { get; init; }

    public static ConditionalFormattingOptions GreaterThan(string formula, IStyle style)
        => CreateThreshold(ConditionalFormattingType.GreaterThan, formula, style);

    public static ConditionalFormattingOptions LessThan(string formula, IStyle style)
        => CreateThreshold(ConditionalFormattingType.LessThan, formula, style);

    public static ConditionalFormattingOptions ContainsText(string text, IStyle style)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            throw new ArgumentException("Conditional formatting text cannot be null or empty.", nameof(text));
        }

        ArgumentNullException.ThrowIfNull(style);

        return new ConditionalFormattingOptions
        {
            Type = ConditionalFormattingType.ContainsText,
            Text = text,
            Style = style
        };
    }

    public static ConditionalFormattingOptions DuplicateValues(IStyle style)
    {
        ArgumentNullException.ThrowIfNull(style);

        return new ConditionalFormattingOptions
        {
            Type = ConditionalFormattingType.DuplicateValues,
            Style = style
        };
    }

    public static ConditionalFormattingOptions TwoColorScale(Color minimumColor, Color maximumColor)
    {
        return new ConditionalFormattingOptions
        {
            Type = ConditionalFormattingType.TwoColorScale,
            MinimumColor = minimumColor,
            MaximumColor = maximumColor
        };
    }

    private static ConditionalFormattingOptions CreateThreshold(ConditionalFormattingType type, string formula, IStyle style)
    {
        if (string.IsNullOrWhiteSpace(formula))
        {
            throw new ArgumentException("Conditional formatting formula cannot be null or empty.", nameof(formula));
        }

        ArgumentNullException.ThrowIfNull(style);

        return new ConditionalFormattingOptions
        {
            Type = type,
            Formula = formula,
            Style = style
        };
    }
}
