using System.Globalization;
using OfficeDocuments.Word.Enums;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.Formatting;

/// <summary>
/// Translates between <see cref="TableFormat"/> and the <c>w:tblPr</c> element behind a table.
/// </summary>
internal static class TableFormatMapper
{
    /// <summary>
    /// Fiftieths of a percent per percent, the unit <c>w:tblW</c> uses for a percentage width.
    /// </summary>
    private const double PercentUnitsPerPercent = 50d;

    /// <summary>
    /// Eighths of a point per point, the unit border widths use.
    /// </summary>
    private const double BorderUnitsPerPoint = 8d;

    /// <summary>
    /// Border width used when borders are requested without an explicit width. Half a point is what
    /// Word applies for a plain single border.
    /// </summary>
    private const double DefaultBorderWidthInPoints = 0.5d;

    /// <summary>
    /// Writes the properties <paramref name="format"/> sets onto <paramref name="table"/>, leaving the
    /// others as they are.
    /// </summary>
    internal static void Apply(WordLib.Table table, TableFormat? format)
    {
        if (format is null || format.IsEmpty)
        {
            return;
        }

        var properties = GetOrCreateProperties(table);

        if (format.StyleId is { } styleId)
        {
            properties.TableStyle = new WordLib.TableStyle { Val = styleId };
        }

        if (format.WidthPercent is { } widthPercent)
        {
            properties.TableWidth = new WordLib.TableWidth
            {
                Type = WordLib.TableWidthUnitValues.Pct,
                Width = ToInvariant((int)Math.Round(widthPercent * PercentUnitsPerPercent, MidpointRounding.AwayFromZero)),
            };
        }

        if (format.Alignment is { } alignment)
        {
            properties.TableJustification = new WordLib.TableJustification { Val = ToOpenXml(alignment) };
        }

        if (format.Borders is { } borders)
        {
            properties.TableBorders = CreateBorders(borders, format);
        }

        if (format.CellPadding is { } cellPadding)
        {
            properties.TableCellMarginDefault = CreateCellMargins(cellPadding);
        }
    }

    /// <summary>
    /// Reads back the table formatting this library models.
    /// </summary>
    internal static TableFormat Read(WordLib.Table table)
    {
        var properties = table.GetFirstChild<WordLib.TableProperties>();
        if (properties is null)
        {
            return new TableFormat();
        }

        return new TableFormat
        {
            StyleId = properties.TableStyle?.Val?.Value,
            WidthPercent = ReadWidthPercent(properties.TableWidth),
            Alignment = FromOpenXml(properties.TableJustification?.Val?.Value),
            Borders = ReadBorders(properties.TableBorders),
            BorderColor = properties.TableBorders?.TopBorder?.Color?.Value,
            BorderWidth = ReadBorderWidth(properties.TableBorders?.TopBorder?.Size?.Value),
            CellPadding = ReadCellPadding(properties.TableCellMarginDefault),
        };
    }

    /// <summary>
    /// Returns the table's properties element, creating it in its required first position.
    /// </summary>
    /// <remarks>
    /// <c>CT_Tbl</c> requires <c>w:tblPr</c> before <c>w:tblGrid</c> and the rows, so this cannot be a
    /// plain append.
    /// </remarks>
    internal static WordLib.TableProperties GetOrCreateProperties(WordLib.Table table)
    {
        var existing = table.GetFirstChild<WordLib.TableProperties>();
        if (existing is not null)
        {
            return existing;
        }

        var properties = new WordLib.TableProperties();
        table.InsertAt(properties, 0);

        return properties;
    }

    private static WordLib.TableBorders CreateBorders(TableBorders borders, TableFormat format)
    {
        var style = borders == Enums.TableBorders.None ? WordLib.BorderValues.None : WordLib.BorderValues.Single;
        var size = (uint)Math.Round(
            (format.BorderWidth ?? DefaultBorderWidthInPoints) * BorderUnitsPerPoint,
            MidpointRounding.AwayFromZero);
        var color = format.BorderColor ?? HexColor.Automatic;

        // Inside borders are written even for Outline, as an explicit "none", so that a table style
        // cannot reinstate the grid lines the caller asked not to have.
        var insideStyle = borders == Enums.TableBorders.All ? style : WordLib.BorderValues.None;

        return new WordLib.TableBorders
        {
            TopBorder = new WordLib.TopBorder { Val = style, Size = size, Color = color },
            LeftBorder = new WordLib.LeftBorder { Val = style, Size = size, Color = color },
            BottomBorder = new WordLib.BottomBorder { Val = style, Size = size, Color = color },
            RightBorder = new WordLib.RightBorder { Val = style, Size = size, Color = color },
            InsideHorizontalBorder = new WordLib.InsideHorizontalBorder { Val = insideStyle, Size = size, Color = color },
            InsideVerticalBorder = new WordLib.InsideVerticalBorder { Val = insideStyle, Size = size, Color = color },
        };
    }

    private static WordLib.TableCellMarginDefault CreateCellMargins(double cellPaddingInPoints)
    {
        var width = Measure.PointsToTwips(cellPaddingInPoints);

        return new WordLib.TableCellMarginDefault
        {
            TopMargin = new WordLib.TopMargin { Width = width, Type = WordLib.TableWidthUnitValues.Dxa },
            TableCellLeftMargin = new WordLib.TableCellLeftMargin { Width = ToInt16(width), Type = WordLib.TableWidthValues.Dxa },
            BottomMargin = new WordLib.BottomMargin { Width = width, Type = WordLib.TableWidthUnitValues.Dxa },
            TableCellRightMargin = new WordLib.TableCellRightMargin { Width = ToInt16(width), Type = WordLib.TableWidthValues.Dxa },
        };
    }

    private static TableBorders? ReadBorders(WordLib.TableBorders? borders)
    {
        var outerStyle = borders?.TopBorder?.Val?.Value;
        if (outerStyle is null)
        {
            return null;
        }

        if (outerStyle == WordLib.BorderValues.None || outerStyle == WordLib.BorderValues.Nil)
        {
            return Enums.TableBorders.None;
        }

        var insideStyle = borders?.InsideHorizontalBorder?.Val?.Value;
        var hasInsideBorders = insideStyle is not null
            && insideStyle != WordLib.BorderValues.None
            && insideStyle != WordLib.BorderValues.Nil;

        return hasInsideBorders ? Enums.TableBorders.All : Enums.TableBorders.Outline;
    }

    private static double? ReadWidthPercent(WordLib.TableWidth? width)
    {
        if (width?.Type?.Value != WordLib.TableWidthUnitValues.Pct || width.Width?.Value is not { } value)
        {
            return null;
        }

        // A percentage width can also be written as "50%" rather than in fiftieths.
        if (value.EndsWith('%')
            && double.TryParse(value[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var percent))
        {
            return percent;
        }

        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var units)
            ? units / PercentUnitsPerPercent
            : null;
    }

    private static double? ReadBorderWidth(uint? size) => size is null ? null : size.Value / BorderUnitsPerPoint;

    private static double? ReadCellPadding(WordLib.TableCellMarginDefault? margins)
    {
        var width = margins?.TopMargin?.Width?.Value;

        return width is not null && double.TryParse(width, NumberStyles.Float, CultureInfo.InvariantCulture, out var twips)
            ? twips / 20d
            : null;
    }

    private static WordLib.TableRowAlignmentValues ToOpenXml(TableAlignment alignment)
    {
        return alignment switch
        {
            TableAlignment.Left => WordLib.TableRowAlignmentValues.Left,
            TableAlignment.Center => WordLib.TableRowAlignmentValues.Center,
            TableAlignment.Right => WordLib.TableRowAlignmentValues.Right,
            _ => throw new ArgumentOutOfRangeException(nameof(alignment), alignment, "Unsupported table alignment."),
        };
    }

    private static TableAlignment? FromOpenXml(WordLib.TableRowAlignmentValues? alignment)
    {
        if (alignment is null)
        {
            return null;
        }

        var value = alignment.Value;

        if (value == WordLib.TableRowAlignmentValues.Left) return TableAlignment.Left;
        if (value == WordLib.TableRowAlignmentValues.Center) return TableAlignment.Center;
        if (value == WordLib.TableRowAlignmentValues.Right) return TableAlignment.Right;

        return null;
    }

    private static string ToInvariant(int value) => value.ToString(CultureInfo.InvariantCulture);

    private static short ToInt16(string twips)
    {
        return short.TryParse(twips, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ? value : (short)0;
    }
}
