using System.Globalization;
using OfficeDocuments.Word.Enums;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.Formatting;

/// <summary>
/// Translates between <see cref="TableCellFormat"/> and the <c>w:tcPr</c> element behind a cell.
/// </summary>
internal static class TableCellFormatMapper
{
    /// <summary>
    /// Fiftieths of a percent per percent, the unit <c>w:tcW</c> uses for a percentage width.
    /// </summary>
    private const double PercentUnitsPerPercent = 50d;

    /// <summary>
    /// Writes the properties <paramref name="format"/> sets onto <paramref name="cell"/>, leaving the
    /// others as they are.
    /// </summary>
    internal static void Apply(WordLib.TableCell cell, TableCellFormat? format)
    {
        if (format is null || format.IsEmpty)
        {
            return;
        }

        var properties = GetOrCreateProperties(cell);

        if (format.WidthPercent is { } widthPercent)
        {
            properties.TableCellWidth = new WordLib.TableCellWidth
            {
                Type = WordLib.TableWidthUnitValues.Pct,
                Width = ((int)Math.Round(widthPercent * PercentUnitsPerPercent, MidpointRounding.AwayFromZero))
                    .ToString(CultureInfo.InvariantCulture),
            };
        }

        if (format.ColumnSpan is { } columnSpan)
        {
            properties.GridSpan = new WordLib.GridSpan { Val = columnSpan };
        }

        if (format.BackgroundColor is { } backgroundColor)
        {
            // A solid fill needs an explicit pattern; without w:val the shading has no effect.
            properties.Shading = new WordLib.Shading
            {
                Val = WordLib.ShadingPatternValues.Clear,
                Color = HexColor.Automatic,
                Fill = backgroundColor,
            };
        }

        if (format.VerticalAlignment is { } verticalAlignment)
        {
            properties.TableCellVerticalAlignment = new WordLib.TableCellVerticalAlignment
            {
                Val = ToOpenXml(verticalAlignment),
            };
        }
    }

    /// <summary>
    /// Reads back the cell formatting this library models.
    /// </summary>
    internal static TableCellFormat Read(WordLib.TableCell cell)
    {
        var properties = cell.GetFirstChild<WordLib.TableCellProperties>();
        if (properties is null)
        {
            return new TableCellFormat();
        }

        return new TableCellFormat
        {
            WidthPercent = ReadWidthPercent(properties.TableCellWidth),
            BackgroundColor = ReadBackgroundColor(properties.Shading),
            VerticalAlignment = FromOpenXml(properties.TableCellVerticalAlignment?.Val?.Value),
            ColumnSpan = properties.GridSpan?.Val?.Value,
        };
    }

    /// <summary>
    /// Returns the cell's properties element, creating it in its required first position.
    /// </summary>
    internal static WordLib.TableCellProperties GetOrCreateProperties(WordLib.TableCell cell)
    {
        var existing = cell.GetFirstChild<WordLib.TableCellProperties>();
        if (existing is not null)
        {
            return existing;
        }

        var properties = new WordLib.TableCellProperties();
        cell.InsertAt(properties, 0);

        return properties;
    }

    private static double? ReadWidthPercent(WordLib.TableCellWidth? width)
    {
        if (width?.Type?.Value != WordLib.TableWidthUnitValues.Pct || width.Width?.Value is not { } value)
        {
            return null;
        }

        if (value.EndsWith('%')
            && double.TryParse(value[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var percent))
        {
            return percent;
        }

        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var units)
            ? units / PercentUnitsPerPercent
            : null;
    }

    private static string? ReadBackgroundColor(WordLib.Shading? shading)
    {
        var fill = shading?.Fill?.Value;

        return string.Equals(fill, HexColor.Automatic, StringComparison.OrdinalIgnoreCase) ? null : fill;
    }

    private static WordLib.TableVerticalAlignmentValues ToOpenXml(CellVerticalAlignment alignment)
    {
        return alignment switch
        {
            CellVerticalAlignment.Top => WordLib.TableVerticalAlignmentValues.Top,
            CellVerticalAlignment.Center => WordLib.TableVerticalAlignmentValues.Center,
            CellVerticalAlignment.Bottom => WordLib.TableVerticalAlignmentValues.Bottom,
            _ => throw new ArgumentOutOfRangeException(nameof(alignment), alignment, "Unsupported cell vertical alignment."),
        };
    }

    private static CellVerticalAlignment? FromOpenXml(WordLib.TableVerticalAlignmentValues? alignment)
    {
        if (alignment is null)
        {
            return null;
        }

        var value = alignment.Value;

        if (value == WordLib.TableVerticalAlignmentValues.Top) return CellVerticalAlignment.Top;
        if (value == WordLib.TableVerticalAlignmentValues.Center) return CellVerticalAlignment.Center;
        if (value == WordLib.TableVerticalAlignmentValues.Bottom) return CellVerticalAlignment.Bottom;

        return null;
    }
}
