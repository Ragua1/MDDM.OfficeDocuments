using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.Formatting;

/// <summary>
/// Definitions for the built-in styles named in <see cref="WordStyleIds"/>.
/// </summary>
/// <remarks>
/// <para>
/// Referencing <c>w:pStyle w:val="Heading1"</c> only changes the look of a paragraph if the document
/// also defines a style with that identifier. Word does not fill the definition in for you, so a
/// library that writes the reference alone produces a document where every heading looks like body
/// text. These definitions exist so that <see cref="ParagraphFormat.StyleId"/> actually renders.
/// </para>
/// <para>
/// <c>w:name</c> carries the built-in name Word recognizes ("heading 1", lower case with a space),
/// which is what makes the style show up in Word's own gallery instead of as a custom style.
/// <c>w:outlineLvl</c> is what puts headings in the navigation pane and in a generated table of
/// contents.
/// </para>
/// </remarks>
internal static class BuiltInStyles
{
    /// <summary>
    /// Look of one built-in style. Sizes and spacing are in points.
    /// </summary>
    private sealed record Definition(
        string StyleId,
        string Name,
        int UiPriority,
        bool IsCharacterStyle = false,
        double? FontSize = null,
        bool Bold = false,
        string? Color = null,
        bool Underline = false,
        double? SpacingBefore = null,
        double? SpacingAfter = null,
        int? OutlineLevel = null,
        bool KeepNext = false,
        bool IsDefault = false);

    private static readonly Dictionary<string, Definition> Definitions = new(StringComparer.Ordinal)
    {
        [WordStyleIds.Normal] = new(WordStyleIds.Normal, "Normal", UiPriority: 1, IsDefault: true),
        [WordStyleIds.Title] = new(WordStyleIds.Title, "Title", UiPriority: 10, FontSize: 28, SpacingAfter: 4, KeepNext: true),
        [WordStyleIds.Subtitle] = new(WordStyleIds.Subtitle, "Subtitle", UiPriority: 11, FontSize: 14, Color: "5A5A5A", SpacingAfter: 16),
        [WordStyleIds.Heading1] = new(WordStyleIds.Heading1, "heading 1", UiPriority: 9, FontSize: 18, Bold: true, Color: "2F5496", SpacingBefore: 12, SpacingAfter: 4, OutlineLevel: 0, KeepNext: true),
        [WordStyleIds.Heading2] = new(WordStyleIds.Heading2, "heading 2", UiPriority: 9, FontSize: 15, Bold: true, Color: "2F5496", SpacingBefore: 10, SpacingAfter: 4, OutlineLevel: 1, KeepNext: true),
        [WordStyleIds.Heading3] = new(WordStyleIds.Heading3, "heading 3", UiPriority: 9, FontSize: 13, Bold: true, Color: "1F3763", SpacingBefore: 10, SpacingAfter: 4, OutlineLevel: 2, KeepNext: true),
        [WordStyleIds.Heading4] = new(WordStyleIds.Heading4, "heading 4", UiPriority: 9, FontSize: 12, Bold: true, Color: "2F5496", SpacingBefore: 8, SpacingAfter: 4, OutlineLevel: 3, KeepNext: true),
        [WordStyleIds.Heading5] = new(WordStyleIds.Heading5, "heading 5", UiPriority: 9, FontSize: 11, Bold: true, Color: "2F5496", SpacingBefore: 8, SpacingAfter: 4, OutlineLevel: 4, KeepNext: true),
        [WordStyleIds.Heading6] = new(WordStyleIds.Heading6, "heading 6", UiPriority: 9, FontSize: 11, Bold: true, Color: "1F3763", SpacingBefore: 8, SpacingAfter: 4, OutlineLevel: 5, KeepNext: true),

        // A character style, not a paragraph style: it formats the runs inside a hyperlink without
        // touching the paragraph that contains them.
        [WordStyleIds.Hyperlink] = new(WordStyleIds.Hyperlink, "Hyperlink", UiPriority: 99, IsCharacterStyle: true, Color: "0563C1", Underline: true),
    };

    /// <summary>
    /// <see langword="true"/> when <paramref name="styleId"/> is a style this library can define.
    /// </summary>
    internal static bool IsKnown(string styleId) => Definitions.ContainsKey(styleId);

    /// <summary>
    /// Builds the style definition element for <paramref name="styleId"/>.
    /// </summary>
    /// <returns>The definition, or <see langword="null"/> when the identifier is not built in.</returns>
    internal static WordLib.Style? TryCreate(string styleId)
    {
        return Definitions.TryGetValue(styleId, out var definition) ? Create(definition) : null;
    }

    private static WordLib.Style Create(Definition definition)
    {
        // Assigning the SDK's typed properties rather than appending children is deliberate: the
        // setters place each element at its schema position, so the strict child sequences of
        // CT_Style, CT_PPr, and CT_RPr hold by construction instead of by our own care.
        var style = new WordLib.Style
        {
            Type = definition.IsCharacterStyle ? WordLib.StyleValues.Character : WordLib.StyleValues.Paragraph,
            StyleId = definition.StyleId,
            StyleName = new WordLib.StyleName { Val = definition.Name },
            UIPriority = new WordLib.UIPriority { Val = definition.UiPriority },
        };

        if (definition.IsCharacterStyle)
        {
            // Word ships the Hyperlink style hidden from the gallery until a document uses it.
            style.SemiHidden = new WordLib.SemiHidden();
            style.UnhideWhenUsed = new WordLib.UnhideWhenUsed();
        }
        else
        {
            style.PrimaryStyle = new WordLib.PrimaryStyle();
        }

        if (definition.IsDefault)
        {
            style.Default = true;
        }
        else if (!definition.IsCharacterStyle)
        {
            style.BasedOn = new WordLib.BasedOn { Val = WordStyleIds.Normal };
            style.NextParagraphStyle = new WordLib.NextParagraphStyle { Val = WordStyleIds.Normal };
        }

        var paragraphProperties = CreateParagraphProperties(definition);
        if (paragraphProperties is not null)
        {
            style.StyleParagraphProperties = paragraphProperties;
        }

        var runProperties = CreateRunProperties(definition);
        if (runProperties is not null)
        {
            style.StyleRunProperties = runProperties;
        }

        return style;
    }

    private static WordLib.StyleParagraphProperties? CreateParagraphProperties(Definition definition)
    {
        if (!definition.KeepNext
            && definition.SpacingBefore is null
            && definition.SpacingAfter is null
            && definition.OutlineLevel is null)
        {
            return null;
        }

        var properties = new WordLib.StyleParagraphProperties();

        if (definition.KeepNext)
        {
            properties.KeepNext = new WordLib.KeepNext();
        }

        if (definition.SpacingBefore is not null || definition.SpacingAfter is not null)
        {
            var spacing = new WordLib.SpacingBetweenLines();
            if (definition.SpacingBefore is { } before)
            {
                spacing.Before = Measure.PointsToTwips(before);
            }

            if (definition.SpacingAfter is { } after)
            {
                spacing.After = Measure.PointsToTwips(after);
            }

            properties.SpacingBetweenLines = spacing;
        }

        if (definition.OutlineLevel is { } outlineLevel)
        {
            properties.OutlineLevel = new WordLib.OutlineLevel { Val = outlineLevel };
        }

        return properties;
    }

    private static WordLib.StyleRunProperties? CreateRunProperties(Definition definition)
    {
        if (!definition.Bold && !definition.Underline && definition.FontSize is null && definition.Color is null)
        {
            return null;
        }

        var properties = new WordLib.StyleRunProperties();

        if (definition.Bold)
        {
            properties.Bold = new WordLib.Bold();
        }

        if (definition.Color is { } color)
        {
            properties.Color = new WordLib.Color { Val = color };
        }

        if (definition.FontSize is { } fontSize)
        {
            var halfPoints = Measure.FontSizeToHalfPoints(fontSize);
            properties.FontSize = new WordLib.FontSize { Val = halfPoints };
            properties.FontSizeComplexScript = new WordLib.FontSizeComplexScript { Val = halfPoints };
        }

        if (definition.Underline)
        {
            properties.Underline = new WordLib.Underline { Val = WordLib.UnderlineValues.Single };
        }

        return properties;
    }
}
