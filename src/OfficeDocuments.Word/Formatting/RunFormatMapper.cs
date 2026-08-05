using System.Globalization;
using OfficeDocuments.Word.Enums;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.Formatting;

/// <summary>
/// Translates between <see cref="TextFormat"/> and the <c>w:rPr</c> element behind a run.
/// </summary>
internal static class RunFormatMapper
{
    /// <summary>
    /// Highlight palette, paired in both directions. A table beats two switch statements here:
    /// seventeen values that must stay in step are easy to get wrong when written out twice.
    /// </summary>
    private static readonly (HighlightColor Model, WordLib.HighlightColorValues OpenXml)[] HighlightMap =
    [
        (HighlightColor.None, WordLib.HighlightColorValues.None),
        (HighlightColor.Yellow, WordLib.HighlightColorValues.Yellow),
        (HighlightColor.Green, WordLib.HighlightColorValues.Green),
        (HighlightColor.Cyan, WordLib.HighlightColorValues.Cyan),
        (HighlightColor.Magenta, WordLib.HighlightColorValues.Magenta),
        (HighlightColor.Blue, WordLib.HighlightColorValues.Blue),
        (HighlightColor.Red, WordLib.HighlightColorValues.Red),
        (HighlightColor.DarkBlue, WordLib.HighlightColorValues.DarkBlue),
        (HighlightColor.DarkCyan, WordLib.HighlightColorValues.DarkCyan),
        (HighlightColor.DarkGreen, WordLib.HighlightColorValues.DarkGreen),
        (HighlightColor.DarkMagenta, WordLib.HighlightColorValues.DarkMagenta),
        (HighlightColor.DarkRed, WordLib.HighlightColorValues.DarkRed),
        (HighlightColor.DarkYellow, WordLib.HighlightColorValues.DarkYellow),
        (HighlightColor.DarkGray, WordLib.HighlightColorValues.DarkGray),
        (HighlightColor.LightGray, WordLib.HighlightColorValues.LightGray),
        (HighlightColor.Black, WordLib.HighlightColorValues.Black),
        (HighlightColor.White, WordLib.HighlightColorValues.White),
    ];

    /// <summary>
    /// Writes the properties <paramref name="format"/> sets onto <paramref name="run"/>, leaving
    /// properties it does not set — including ones this library does not model — untouched.
    /// </summary>
    /// <param name="run">Run element to format.</param>
    /// <param name="format">Properties to write. Ignored when empty.</param>
    /// <param name="ensureStyle">
    /// Callback that makes a referenced character style exist in the document.
    /// </param>
    internal static void Apply(WordLib.Run run, TextFormat? format, Action<string> ensureStyle)
    {
        if (format is null || format.IsEmpty)
        {
            return;
        }

        // The SDK's typed setters place each element at its position in the CT_RPr sequence, so the
        // order requirement is satisfied by construction rather than by remembering it here.
        var properties = run.RunProperties ??= new WordLib.RunProperties();

        if (format.StyleId is { } styleId)
        {
            ensureStyle(styleId);
            properties.RunStyle = new WordLib.RunStyle { Val = styleId };
        }

        if (format.Bold is { } bold)
        {
            properties.Bold = bold ? new WordLib.Bold() : new WordLib.Bold { Val = false };
        }

        if (format.Italic is { } italic)
        {
            properties.Italic = italic ? new WordLib.Italic() : new WordLib.Italic { Val = false };
        }

        if (format.Strikethrough is { } strikethrough)
        {
            properties.Strike = strikethrough ? new WordLib.Strike() : new WordLib.Strike { Val = false };
        }

        if (format.AllCaps is { } allCaps)
        {
            properties.Caps = allCaps ? new WordLib.Caps() : new WordLib.Caps { Val = false };
        }

        if (format.SmallCaps is { } smallCaps)
        {
            properties.SmallCaps = smallCaps ? new WordLib.SmallCaps() : new WordLib.SmallCaps { Val = false };
        }

        if (format.Underline is { } underline)
        {
            properties.Underline = new WordLib.Underline { Val = ToOpenXml(underline) };
        }

        if (format.Highlight is { } highlight)
        {
            properties.Highlight = new WordLib.Highlight { Val = ToOpenXml(highlight) };
        }

        if (format.VerticalPosition is { } verticalPosition)
        {
            properties.VerticalTextAlignment = new WordLib.VerticalTextAlignment { Val = ToOpenXml(verticalPosition) };
        }

        if (format.FontName is { } fontName)
        {
            var fonts = properties.RunFonts ??= new WordLib.RunFonts();
            fonts.Ascii = fontName;
            fonts.HighAnsi = fontName;
            fonts.ComplexScript = fontName;
        }

        if (format.FontSize is { } fontSize)
        {
            var halfPoints = Measure.FontSizeToHalfPoints(fontSize);
            properties.FontSize = new WordLib.FontSize { Val = halfPoints };
            properties.FontSizeComplexScript = new WordLib.FontSizeComplexScript { Val = halfPoints };
        }

        if (format.Color is { } color)
        {
            properties.Color = new WordLib.Color { Val = color };
        }
    }

    /// <summary>
    /// Reads back the direct formatting this library models. Properties the run does not set, and
    /// properties outside <see cref="TextFormat"/>, come back as <see langword="null"/>.
    /// </summary>
    internal static TextFormat Read(WordLib.Run run)
    {
        var properties = run.RunProperties;
        if (properties is null)
        {
            return new TextFormat();
        }

        return new TextFormat
        {
            StyleId = properties.RunStyle?.Val?.Value,
            Bold = ReadToggle(properties.Bold),
            Italic = ReadToggle(properties.Italic),
            Strikethrough = ReadToggle(properties.Strike),
            AllCaps = ReadToggle(properties.Caps),
            SmallCaps = ReadToggle(properties.SmallCaps),
            Underline = FromOpenXml(properties.Underline?.Val?.Value),
            Highlight = FromOpenXml(properties.Highlight?.Val?.Value),
            VerticalPosition = FromOpenXml(properties.VerticalTextAlignment?.Val?.Value),
            FontName = properties.RunFonts?.Ascii?.Value,
            FontSize = ReadFontSize(properties.FontSize?.Val?.Value),
            Color = properties.Color?.Val?.Value,
        };
    }

    /// <summary>
    /// Reads a toggle property: absent means unset, present without <c>w:val</c> means "on", and
    /// present with a value means what the value says.
    /// </summary>
    private static bool? ReadToggle(WordLib.OnOffType? element)
    {
        if (element is null)
        {
            return null;
        }

        return element.Val is null || element.Val.Value;
    }

    private static double? ReadFontSize(string? halfPoints)
    {
        if (halfPoints is null || !int.TryParse(halfPoints, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
        {
            return null;
        }

        return parsed / 2d;
    }

    private static WordLib.HighlightColorValues ToOpenXml(HighlightColor highlight)
    {
        foreach (var (model, openXml) in HighlightMap)
        {
            if (model == highlight)
            {
                return openXml;
            }
        }

        throw new ArgumentOutOfRangeException(nameof(highlight), highlight, "Unsupported highlight color.");
    }

    private static HighlightColor? FromOpenXml(WordLib.HighlightColorValues? highlight)
    {
        if (highlight is null)
        {
            return null;
        }

        foreach (var (model, openXml) in HighlightMap)
        {
            if (openXml == highlight.Value)
            {
                return model;
            }
        }

        return null;
    }

    private static WordLib.VerticalPositionValues ToOpenXml(TextVerticalPosition verticalPosition)
    {
        return verticalPosition switch
        {
            TextVerticalPosition.Baseline => WordLib.VerticalPositionValues.Baseline,
            TextVerticalPosition.Superscript => WordLib.VerticalPositionValues.Superscript,
            TextVerticalPosition.Subscript => WordLib.VerticalPositionValues.Subscript,
            _ => throw new ArgumentOutOfRangeException(nameof(verticalPosition), verticalPosition, "Unsupported vertical text position."),
        };
    }

    private static TextVerticalPosition? FromOpenXml(WordLib.VerticalPositionValues? verticalPosition)
    {
        if (verticalPosition is null)
        {
            return null;
        }

        var value = verticalPosition.Value;

        if (value == WordLib.VerticalPositionValues.Baseline) return TextVerticalPosition.Baseline;
        if (value == WordLib.VerticalPositionValues.Superscript) return TextVerticalPosition.Superscript;
        if (value == WordLib.VerticalPositionValues.Subscript) return TextVerticalPosition.Subscript;

        return null;
    }

    private static WordLib.UnderlineValues ToOpenXml(UnderlineType underline)
    {
        return underline switch
        {
            UnderlineType.None => WordLib.UnderlineValues.None,
            UnderlineType.Single => WordLib.UnderlineValues.Single,
            UnderlineType.Double => WordLib.UnderlineValues.Double,
            UnderlineType.Thick => WordLib.UnderlineValues.Thick,
            UnderlineType.Dotted => WordLib.UnderlineValues.Dotted,
            UnderlineType.Dash => WordLib.UnderlineValues.Dash,
            UnderlineType.Wave => WordLib.UnderlineValues.Wave,
            UnderlineType.Words => WordLib.UnderlineValues.Words,
            _ => throw new ArgumentOutOfRangeException(nameof(underline), underline, "Unsupported underline style."),
        };
    }

    private static UnderlineType? FromOpenXml(WordLib.UnderlineValues? underline)
    {
        if (underline is null)
        {
            return null;
        }

        var value = underline.Value;

        if (value == WordLib.UnderlineValues.None) return UnderlineType.None;
        if (value == WordLib.UnderlineValues.Single) return UnderlineType.Single;
        if (value == WordLib.UnderlineValues.Double) return UnderlineType.Double;
        if (value == WordLib.UnderlineValues.Thick) return UnderlineType.Thick;
        if (value == WordLib.UnderlineValues.Dotted) return UnderlineType.Dotted;
        if (value == WordLib.UnderlineValues.Dash) return UnderlineType.Dash;
        if (value == WordLib.UnderlineValues.Wave) return UnderlineType.Wave;
        if (value == WordLib.UnderlineValues.Words) return UnderlineType.Words;

        // A style this library does not model, such as dotDash. Reporting null keeps "unset" and
        // "unmodelled" indistinguishable, which is honest: we cannot round-trip it through TextFormat.
        return null;
    }
}
