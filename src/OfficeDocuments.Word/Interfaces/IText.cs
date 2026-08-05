namespace OfficeDocuments.Word.Interfaces;

/// <summary>
/// A single text element inside a run, as returned by
/// <see cref="IParagraph.GetTextElements"/>.
/// </summary>
/// <remarks>
/// This is the element-level view, one step below <see cref="IRun"/>: a run can hold several text
/// elements separated by breaks or tabs. Prefer <see cref="IRun.Text"/> for whole-run edits and
/// <see cref="IParagraph.GetTexts"/> for reading; reach for this only when the individual element
/// matters.
/// </remarks>
public interface IText
{
    /// <summary>
    /// Gets or sets the element's text. Whitespace is preserved as written.
    /// </summary>
    string TextValue { get; set; }
}
