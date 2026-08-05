namespace OfficeDocuments.Word.Formatting;

/// <summary>
/// Identifiers of the built-in paragraph styles this library can define on demand.
/// </summary>
/// <remarks>
/// Assign one to <see cref="ParagraphFormat.StyleId"/>. A referenced style has to exist in the
/// document's style definitions to render, so the library adds a definition the first time a style
/// is used. Any other identifier is written through untouched, which is what a document created from
/// a template with its own styles needs.
/// </remarks>
public static class WordStyleIds
{
    /// <summary>The default body style every other style is based on.</summary>
    public const string Normal = "Normal";

    /// <summary>Document title.</summary>
    public const string Title = "Title";

    /// <summary>Secondary line below the title.</summary>
    public const string Subtitle = "Subtitle";

    /// <summary>Top-level heading; outline level 1.</summary>
    public const string Heading1 = "Heading1";

    /// <summary>Second-level heading.</summary>
    public const string Heading2 = "Heading2";

    /// <summary>Third-level heading.</summary>
    public const string Heading3 = "Heading3";

    /// <summary>Fourth-level heading.</summary>
    public const string Heading4 = "Heading4";

    /// <summary>Fifth-level heading.</summary>
    public const string Heading5 = "Heading5";

    /// <summary>Sixth-level heading.</summary>
    public const string Heading6 = "Heading6";

    /// <summary>
    /// Character style applied to the text of a hyperlink: blue and underlined.
    /// </summary>
    /// <remarks>
    /// Unlike the others this is a character style, so it formats runs rather than paragraphs. It is
    /// applied automatically by <see cref="Interfaces.IParagraph.AddHyperlink(string, string, TextFormat?)"/>.
    /// </remarks>
    public const string Hyperlink = "Hyperlink";

    /// <summary>
    /// Returns the built-in heading identifier for <paramref name="level"/> (1 to 6).
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The level is outside 1 to 6.</exception>
    public static string Heading(int level)
    {
        return level switch
        {
            1 => Heading1,
            2 => Heading2,
            3 => Heading3,
            4 => Heading4,
            5 => Heading5,
            6 => Heading6,
            _ => throw new ArgumentOutOfRangeException(nameof(level), level, "Heading level must be between 1 and 6."),
        };
    }
}
