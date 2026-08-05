using System.Text;
using DocumentFormat.OpenXml;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.DataClasses;

/// <summary>
/// Writes and reads the textual content of runs, hiding two WordprocessingML details that callers
/// otherwise have to know about.
/// </summary>
/// <remarks>
/// <para>
/// <b>Whitespace.</b> An XML processor is free to collapse whitespace, so <c>&lt;w:t&gt;Total: &lt;/w:t&gt;</c>
/// loses its trailing space and the sentence closes up. Keeping it requires <c>xml:space="preserve"</c>,
/// which this writer adds whenever the text starts or ends with whitespace.
/// </para>
/// <para>
/// <b>Line breaks.</b> A newline inside <c>w:t</c> is just whitespace to Word; the break element is
/// <c>w:br</c>. Text containing <c>\n</c> is therefore split into alternating text and break
/// elements, so that a caller can pass a multi-line string and get multiple lines.
/// </para>
/// </remarks>
internal static class RunContent
{
    private static readonly char[] NewLineCharacters = ['\n', '\r'];

    /// <summary>
    /// Appends <paramref name="text"/> to <paramref name="run"/>, translating newlines into breaks.
    /// </summary>
    internal static void Append(WordLib.Run run, string text)
    {
        foreach (var element in CreateContent(text, keepEmpty: true))
        {
            run.AppendChild(element);
        }
    }

    /// <summary>
    /// Builds the content elements that represent <paramref name="text"/>, without attaching them.
    /// </summary>
    /// <remarks>
    /// Separated from <see cref="Append"/> so that text replacement can splice the same elements in at
    /// a position rather than at the end of a run. Both paths then produce identical markup, which is
    /// the point: a caller should not be able to tell whether text was authored or replaced in.
    /// </remarks>
    /// <param name="text">Text to represent. Newlines become <c>w:br</c>.</param>
    /// <param name="keepEmpty">
    /// <see langword="true"/> to return one empty <c>w:t</c> for empty text, which is what a new run
    /// needs so that it reads back as empty rather than as absent. <see langword="false"/> to return
    /// nothing, which is what replacing text with nothing needs.
    /// </param>
    internal static List<OpenXmlElement> CreateContent(string text, bool keepEmpty)
    {
        if (text.Length == 0)
        {
            return keepEmpty ? [CreateText(string.Empty)] : [];
        }

        if (text.IndexOfAny(NewLineCharacters) < 0)
        {
            return [CreateText(text)];
        }

        var elements = new List<OpenXmlElement>();
        var lines = text.Split(["\r\n", "\n", "\r"], StringSplitOptions.None);
        for (var index = 0; index < lines.Length; index++)
        {
            if (index > 0)
            {
                elements.Add(new WordLib.Break());
            }

            if (lines[index].Length > 0)
            {
                elements.Add(CreateText(lines[index]));
            }
        }

        return elements;
    }

    /// <summary>
    /// Replaces the textual content of <paramref name="run"/>, leaving its formatting in place.
    /// </summary>
    internal static void Replace(WordLib.Run run, string text)
    {
        foreach (var element in run.ChildElements.Where(IsContent).ToList())
        {
            element.Remove();
        }

        Append(run, text);
    }

    /// <summary>
    /// Reads the text of every run under <paramref name="element"/> in document order.
    /// </summary>
    /// <remarks>
    /// Walking descendants rather than direct children means runs nested in a container — a hyperlink,
    /// for instance — are read too. Breaks become <c>\n</c> and tabs <c>\t</c>, so the result reflects
    /// the document's line structure instead of silently running lines together.
    /// </remarks>
    internal static string Read(OpenXmlElement element)
    {
        var builder = new StringBuilder();

        foreach (var (_, text) in Enumerate(element))
        {
            builder.Append(text);
        }

        return builder.ToString();
    }

    /// <summary>
    /// Walks the text-bearing elements under <paramref name="container"/> in document order, pairing
    /// each with the text it contributes.
    /// </summary>
    /// <remarks>
    /// The single definition of what a document's text is and which element produced which characters.
    /// <see cref="Read"/> and the offset arithmetic in <see cref="TextReplacer"/> both derive from this
    /// walk rather than each implementing it, because a search that indexes into one string and edits
    /// through another has to agree with itself exactly — an <c>IndexOf</c> result is meaningless
    /// against a different notion of the same text.
    /// </remarks>
    internal static IEnumerable<(OpenXmlElement Element, string Text)> Enumerate(OpenXmlElement container)
    {
        foreach (var descendant in container.Descendants())
        {
            switch (descendant)
            {
                case WordLib.Text text:
                    yield return (text, text.Text ?? string.Empty);
                    break;
                case WordLib.Break:
                    yield return (descendant, "\n");
                    break;
                case WordLib.TabChar:
                    yield return (descendant, "\t");
                    break;
            }
        }
    }

    private static bool IsContent(OpenXmlElement element)
    {
        return element is WordLib.Text or WordLib.Break or WordLib.TabChar;
    }

    private static WordLib.Text CreateText(string value)
    {
        var text = new WordLib.Text(value);

        if (value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1])))
        {
            text.Space = SpaceProcessingModeValues.Preserve;
        }

        return text;
    }
}
