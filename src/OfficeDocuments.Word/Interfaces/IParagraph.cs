using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;

namespace OfficeDocuments.Word.Interfaces;

/// <summary>
/// A paragraph: the block that holds runs of text and carries paragraph-level formatting.
/// </summary>
public interface IParagraph
{
    /// <summary>
    /// The runs of this paragraph, in document order.
    /// </summary>
    IReadOnlyList<IRun> Runs { get; }

    /// <summary>
    /// The direct paragraph formatting this library models. Properties the paragraph does not set
    /// come back as <see langword="null"/>.
    /// </summary>
    ParagraphFormat Format { get; }

    /// <summary>
    /// Applies the properties <paramref name="format"/> sets, leaving the others as they are.
    /// </summary>
    /// <remarks>
    /// A <see cref="ParagraphFormat.StyleId"/> naming a built-in style also adds that style's
    /// definition to the document, so the paragraph actually looks like the style.
    /// </remarks>
    /// <param name="format">Formatting to apply.</param>
    /// <returns>This paragraph, for chaining.</returns>
    IParagraph ApplyFormat(ParagraphFormat format);

    /// <summary>
    /// Appends unformatted text.
    /// </summary>
    /// <param name="text">Text to append. Newlines become line breaks.</param>
    /// <returns>This paragraph, for chaining.</returns>
    IParagraph AddText(string text);

    /// <summary>
    /// Appends formatted text.
    /// </summary>
    /// <param name="text">Text to append. Newlines become line breaks.</param>
    /// <param name="format">Character formatting, or <see langword="null"/> to inherit the style's.</param>
    /// <returns>This paragraph, for chaining.</returns>
    IParagraph AddText(string text, TextFormat? format);

    /// <summary>
    /// Appends a run and returns it, for when the run itself is needed rather than the paragraph.
    /// </summary>
    /// <param name="text">Text of the run. Newlines become line breaks.</param>
    /// <param name="format">Character formatting, or <see langword="null"/> to inherit the style's.</param>
    /// <returns>The new run.</returns>
    IRun AddRun(string text, TextFormat? format = null);

    /// <summary>
    /// Appends a break.
    /// </summary>
    /// <param name="type">Kind of break.</param>
    /// <returns>This paragraph, for chaining.</returns>
    IParagraph AddBreak(BreakType type);

    /// <summary>
    /// Appends a hyperlink to an external target.
    /// </summary>
    /// <remarks>
    /// The link's text is styled with the built-in <see cref="WordStyleIds.Hyperlink"/> character
    /// style, which the library defines in the document on first use, so the link looks like one.
    /// </remarks>
    /// <param name="text">Text shown to the reader.</param>
    /// <param name="url">Absolute target, for example <c>https://example.com</c> or a <c>mailto:</c> address.</param>
    /// <param name="format">Extra character formatting layered over the hyperlink style.</param>
    /// <returns>This paragraph, for chaining.</returns>
    /// <exception cref="ArgumentException">The URL is empty or not an absolute URI.</exception>
    IParagraph AddHyperlink(string text, string url, TextFormat? format = null);

    /// <summary>
    /// Appends an inline image read from <paramref name="content"/>, inferring its format.
    /// </summary>
    /// <param name="content">Image bytes. Must be seekable so the format and size can be read.</param>
    /// <param name="size">Rendered size, or <see langword="null"/> for the image's own size.</param>
    /// <param name="description">Alternative text for accessibility.</param>
    /// <returns>This paragraph, for chaining.</returns>
    /// <exception cref="ArgumentException">The image format could not be determined.</exception>
    IParagraph AddImage(Stream content, ImageSize? size = null, string? description = null);

    /// <summary>
    /// Appends an inline image of a known format.
    /// </summary>
    /// <remarks>
    /// The overload to use for a stream that cannot seek, which also requires an
    /// <see cref="ImageSize.Exact"/> size because the image's own dimensions cannot be read.
    /// </remarks>
    /// <param name="content">Image bytes.</param>
    /// <param name="imageType">Format of the image.</param>
    /// <param name="size">Rendered size, or <see langword="null"/> for the image's own size.</param>
    /// <param name="description">Alternative text for accessibility.</param>
    /// <returns>This paragraph, for chaining.</returns>
    IParagraph AddImage(Stream content, ImageType imageType, ImageSize? size = null, string? description = null);

    /// <summary>
    /// Appends an inline image read from a file, inferring its format from the extension.
    /// </summary>
    /// <param name="filePath">Path of the image file.</param>
    /// <param name="size">Rendered size, or <see langword="null"/> for the image's own size.</param>
    /// <param name="description">Alternative text for accessibility.</param>
    /// <returns>This paragraph, for chaining.</returns>
    /// <exception cref="ArgumentException">The extension names no supported image type.</exception>
    /// <exception cref="FileNotFoundException">The file does not exist.</exception>
    IParagraph AddImage(string filePath, ImageSize? size = null, string? description = null);

    /// <summary>
    /// Replaces everything the paragraph contains with <paramref name="text"/>, keeping its formatting.
    /// </summary>
    /// <remarks>
    /// The paragraph's own formatting and style survive; its runs, hyperlinks, and images do not. Use
    /// <see cref="ReplaceText"/> instead to change part of the text and keep the rest.
    /// </remarks>
    /// <param name="text">New content. Newlines become line breaks.</param>
    /// <param name="format">Character formatting for the new text, or <see langword="null"/> to inherit.</param>
    /// <returns>This paragraph, for chaining.</returns>
    IParagraph SetText(string text, TextFormat? format = null);

    /// <summary>
    /// Replaces every occurrence of <paramref name="oldValue"/> in this paragraph's text.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Works on the paragraph's text as a whole, so a match is found even when Word has split it across
    /// several runs — which it routinely has, because spell-check state and revision tracking start new
    /// runs mid-word. A placeholder that looks like one word in Word is very often three runs on disk.
    /// </para>
    /// <para>
    /// The replacement takes the character formatting of the run where the match starts. A match cannot
    /// span two paragraphs.
    /// </para>
    /// </remarks>
    /// <param name="oldValue">Text to find.</param>
    /// <param name="newValue">Replacement text. Newlines become line breaks; empty text deletes.</param>
    /// <param name="comparison">How to compare. Ordinal by default.</param>
    /// <returns>The number of occurrences replaced.</returns>
    /// <exception cref="ArgumentException"><paramref name="oldValue"/> is empty.</exception>
    int ReplaceText(string oldValue, string newValue, StringComparison comparison = StringComparison.Ordinal);

    /// <summary>
    /// The individual text elements of this paragraph.
    /// </summary>
    IEnumerable<IText> GetTextElements();

    /// <summary>
    /// The paragraph's text, with line breaks as <c>\n</c> and tabs as <c>\t</c>.
    /// </summary>
    string GetTexts();
}
