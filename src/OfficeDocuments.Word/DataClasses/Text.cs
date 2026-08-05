using OfficeDocuments.Word.Interfaces;

namespace OfficeDocuments.Word.DataClasses;

/// <summary>
/// Wraps a single <c>w:t</c> element. See <see cref="IText"/> for when to use it.
/// </summary>
public class Text : IText
{
    /// <inheritdoc />
    public string TextValue
    {
        get => Element.Text;
        set => Element.Text = value;
    }

    internal DocumentFormat.OpenXml.Wordprocessing.Text Element { get; }

    /// <summary>
    /// Wraps an existing <c>w:t</c> element.
    /// </summary>
    /// <param name="element">The element to wrap.</param>
    public Text(DocumentFormat.OpenXml.Wordprocessing.Text element)
    {
        Element = element;
    }
}
