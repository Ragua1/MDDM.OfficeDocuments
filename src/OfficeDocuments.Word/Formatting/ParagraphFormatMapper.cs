using System.Globalization;
using OfficeDocuments.Word.Enums;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.Formatting;

/// <summary>
/// Translates between <see cref="ParagraphFormat"/> and the <c>w:pPr</c> element behind a paragraph.
/// </summary>
internal static class ParagraphFormatMapper
{
    /// <summary>
    /// Twips in one line of single spacing, as <c>w:spacing/@w:line</c> defines it.
    /// </summary>
    private const double TwipsPerLine = 240d;

    /// <summary>
    /// Writes the properties <paramref name="format"/> sets onto <paramref name="paragraph"/>,
    /// leaving properties it does not set untouched.
    /// </summary>
    /// <param name="paragraph">Paragraph element to format.</param>
    /// <param name="format">Properties to write. Ignored when empty.</param>
    /// <param name="ensureStyle">
    /// Callback that makes a referenced style exist in the document. Passing the callbacks in rather
    /// than the document keeps this mapper free of document-level concerns.
    /// </param>
    /// <param name="resolveListNumbering">
    /// Callback that returns the numbering identifier for a list style, defining the numbering in the
    /// document if needed.
    /// </param>
    internal static void Apply(
        WordLib.Paragraph paragraph,
        ParagraphFormat? format,
        Action<string> ensureStyle,
        Func<ListStyle, int> resolveListNumbering)
    {
        if (format is null || format.IsEmpty)
        {
            return;
        }

        // As in the run mapper, typed setters keep the strict CT_PPr child order correct for us.
        var properties = paragraph.ParagraphProperties ??= new WordLib.ParagraphProperties();

        if (format.StyleId is { } styleId)
        {
            ensureStyle(styleId);
            properties.ParagraphStyleId = new WordLib.ParagraphStyleId { Val = styleId };
        }

        if (format.KeepWithNext is { } keepWithNext)
        {
            properties.KeepNext = keepWithNext ? new WordLib.KeepNext() : new WordLib.KeepNext { Val = false };
        }

        if (format.KeepLines is { } keepLines)
        {
            properties.KeepLines = keepLines ? new WordLib.KeepLines() : new WordLib.KeepLines { Val = false };
        }

        if (format.PageBreakBefore is { } pageBreakBefore)
        {
            properties.PageBreakBefore = pageBreakBefore
                ? new WordLib.PageBreakBefore()
                : new WordLib.PageBreakBefore { Val = false };
        }

        ApplyListNumbering(properties, format, resolveListNumbering);

        if (format.Alignment is { } alignment)
        {
            properties.Justification = new WordLib.Justification { Val = ToOpenXml(alignment) };
        }

        ApplySpacing(properties, format);
        ApplyIndentation(properties, format);
    }

    /// <summary>
    /// Reads back the paragraph formatting this library models.
    /// </summary>
    /// <param name="paragraph">Paragraph element to read.</param>
    /// <param name="resolveListStyle">
    /// Callback that classifies a numbering identifier as a bullet or numbered list, or returns
    /// <see langword="null"/> when it is neither.
    /// </param>
    internal static ParagraphFormat Read(WordLib.Paragraph paragraph, Func<int, ListStyle?> resolveListStyle)
    {
        var properties = paragraph.ParagraphProperties;
        if (properties is null)
        {
            return new ParagraphFormat();
        }

        var spacing = properties.SpacingBetweenLines;
        var indentation = properties.Indentation;
        var numbering = properties.NumberingProperties;

        return new ParagraphFormat
        {
            StyleId = properties.ParagraphStyleId?.Val?.Value,
            KeepWithNext = ReadToggle(properties.KeepNext),
            KeepLines = ReadToggle(properties.KeepLines),
            PageBreakBefore = ReadToggle(properties.PageBreakBefore),
            ListStyle = ReadListStyle(numbering, resolveListStyle),
            ListLevel = numbering?.NumberingLevelReference?.Val?.Value,
            Alignment = FromOpenXml(properties.Justification?.Val?.Value),
            SpacingBefore = TwipsToPoints(spacing?.Before?.Value),
            SpacingAfter = TwipsToPoints(spacing?.After?.Value),
            LineSpacing = ReadLineSpacing(spacing),
            IndentLeft = TwipsToPoints(indentation?.Left?.Value),
            IndentRight = TwipsToPoints(indentation?.Right?.Value),
            IndentFirstLine = ReadFirstLineIndent(indentation),
        };
    }

    private static void ApplyListNumbering(
        WordLib.ParagraphProperties properties,
        ParagraphFormat format,
        Func<ListStyle, int> resolveListNumbering)
    {
        if (format.ListStyle is null && format.ListLevel is null)
        {
            return;
        }

        var numbering = properties.NumberingProperties ??= new WordLib.NumberingProperties();

        if (format.ListLevel is { } listLevel)
        {
            numbering.NumberingLevelReference = new WordLib.NumberingLevelReference { Val = listLevel };
        }

        if (format.ListStyle is { } listStyle)
        {
            // ListStyle.None resolves to numbering id 0, which the format reserves for "no list".
            numbering.NumberingId = new WordLib.NumberingId { Val = resolveListNumbering(listStyle) };
        }
    }

    private static ListStyle? ReadListStyle(WordLib.NumberingProperties? numbering, Func<int, ListStyle?> resolveListStyle)
    {
        if (numbering?.NumberingId?.Val?.Value is not { } numberingId)
        {
            return null;
        }

        return numberingId == ListNumbering.NoNumberingId ? ListStyle.None : resolveListStyle(numberingId);
    }

    /// <summary>
    /// Reads a toggle property: absent means unset, present without <c>w:val</c> means "on".
    /// </summary>
    private static bool? ReadToggle(WordLib.OnOffType? element)
    {
        if (element is null)
        {
            return null;
        }

        return element.Val is null || element.Val.Value;
    }

    private static void ApplySpacing(WordLib.ParagraphProperties properties, ParagraphFormat format)
    {
        if (format.SpacingBefore is null && format.SpacingAfter is null && format.LineSpacing is null)
        {
            return;
        }

        var spacing = properties.SpacingBetweenLines ??= new WordLib.SpacingBetweenLines();

        if (format.SpacingBefore is { } before)
        {
            spacing.Before = Measure.PointsToTwips(before);

            // A style that switches auto-spacing on wins over an explicit value, so an explicit
            // value has to switch it off or the requested spacing is quietly ignored by Word.
            spacing.BeforeAutoSpacing = false;
        }

        if (format.SpacingAfter is { } after)
        {
            spacing.After = Measure.PointsToTwips(after);
            spacing.AfterAutoSpacing = false;
        }

        if (format.LineSpacing is { } lineSpacing)
        {
            spacing.Line = Measure.LineSpacingToTwips(lineSpacing);
            spacing.LineRule = WordLib.LineSpacingRuleValues.Auto;
        }
    }

    private static void ApplyIndentation(WordLib.ParagraphProperties properties, ParagraphFormat format)
    {
        if (format.IndentLeft is null && format.IndentRight is null && format.IndentFirstLine is null)
        {
            return;
        }

        var indentation = properties.Indentation ??= new WordLib.Indentation();

        if (format.IndentLeft is { } left)
        {
            indentation.Left = Measure.PointsToTwips(left);
        }

        if (format.IndentRight is { } right)
        {
            indentation.Right = Measure.PointsToTwips(right);
        }

        if (format.IndentFirstLine is { } firstLine)
        {
            // w:ind expresses the two directions as different attributes, and setting both is
            // contradictory, so the sign picks one and clears the other.
            if (firstLine < 0d)
            {
                indentation.Hanging = Measure.PointsToTwips(-firstLine);
                indentation.FirstLine = null;
            }
            else
            {
                indentation.FirstLine = Measure.PointsToTwips(firstLine);
                indentation.Hanging = null;
            }
        }
    }

    private static double? ReadLineSpacing(WordLib.SpacingBetweenLines? spacing)
    {
        var line = spacing?.Line?.Value;
        if (line is null)
        {
            return null;
        }

        // "atLeast" and "exactly" store an absolute height rather than a multiple, which
        // ParagraphFormat.LineSpacing cannot express.
        var rule = spacing?.LineRule?.Value;
        if (rule is not null && rule != WordLib.LineSpacingRuleValues.Auto)
        {
            return null;
        }

        return TryParseTwips(line, out var twips) ? twips / TwipsPerLine : null;
    }

    private static double? ReadFirstLineIndent(WordLib.Indentation? indentation)
    {
        if (indentation?.Hanging?.Value is { } hanging)
        {
            return TryParseTwips(hanging, out var twips) ? -(twips / 20d) : null;
        }

        return TwipsToPoints(indentation?.FirstLine?.Value);
    }

    private static double? TwipsToPoints(string? value)
    {
        return TryParseTwips(value, out var twips) ? twips / 20d : null;
    }

    private static bool TryParseTwips(string? value, out double twips)
    {
        twips = 0d;

        return value is not null
            && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out twips);
    }

    private static WordLib.JustificationValues ToOpenXml(ParagraphAlignment alignment)
    {
        return alignment switch
        {
            ParagraphAlignment.Left => WordLib.JustificationValues.Left,
            ParagraphAlignment.Center => WordLib.JustificationValues.Center,
            ParagraphAlignment.Right => WordLib.JustificationValues.Right,
            ParagraphAlignment.Justify => WordLib.JustificationValues.Both,
            _ => throw new ArgumentOutOfRangeException(nameof(alignment), alignment, "Unsupported paragraph alignment."),
        };
    }

    private static ParagraphAlignment? FromOpenXml(WordLib.JustificationValues? justification)
    {
        if (justification is null)
        {
            return null;
        }

        var value = justification.Value;

        if (value == WordLib.JustificationValues.Left) return ParagraphAlignment.Left;
        if (value == WordLib.JustificationValues.Center) return ParagraphAlignment.Center;
        if (value == WordLib.JustificationValues.Right) return ParagraphAlignment.Right;
        if (value == WordLib.JustificationValues.Both) return ParagraphAlignment.Justify;

        // Values such as "distribute" have no ParagraphAlignment equivalent.
        return null;
    }
}
