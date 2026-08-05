using OfficeDocuments.Word.Formatting;

namespace OfficeDocuments.Word.Interfaces;

/// <summary>
/// A run: a span of text inside a paragraph that shares one set of character formatting.
/// </summary>
public interface IRun
{
    /// <summary>
    /// The run's text. Reading joins the run's text content, rendering line breaks as <c>\n</c> and
    /// tabs as <c>\t</c>; assigning replaces that content and keeps the formatting.
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// The direct character formatting this library models. Properties the run does not set come back
    /// as <see langword="null"/>, which means "inherited from the paragraph's style".
    /// </summary>
    TextFormat Format { get; }

    /// <summary>
    /// Applies the properties <paramref name="format"/> sets, leaving the others as they are.
    /// </summary>
    /// <param name="format">Formatting to apply.</param>
    /// <returns>This run, for chaining.</returns>
    IRun ApplyFormat(TextFormat format);
}
