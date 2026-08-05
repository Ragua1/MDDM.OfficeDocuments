using DocumentFormat.OpenXml;
using OfficeDocuments.Word.Interfaces;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.DataClasses;

/// <summary>
/// The document body: the ordered block content of a <c>.docx</c>.
/// </summary>
public class Body : BlockContainer, IBody
{
    internal new WordLib.Body Element { get; }

    internal Body(WordLib.Body element, DocumentContext context) : base(element, context)
    {
        Element = element;
    }

    /// <summary>
    /// Returns the body's section properties, adding them if the document has none.
    /// </summary>
    /// <remarks>
    /// These carry the page setup and the header and footer references, and the schema requires them to
    /// be the last child of <c>w:body</c>.
    /// </remarks>
    internal WordLib.SectionProperties GetOrCreateSectionProperties()
    {
        var existing = Element.GetFirstChild<WordLib.SectionProperties>();
        if (existing is not null)
        {
            return existing;
        }

        var sectionProperties = new WordLib.SectionProperties();
        Element.AppendChild(sectionProperties);

        return sectionProperties;
    }

    /// <summary>
    /// Appends a block-level element, keeping the body's trailing section properties last.
    /// </summary>
    /// <remarks>
    /// <c>CT_Body</c> is <c>(block-level content)*, sectPr?</c>, so <c>w:sectPr</c> has to remain the
    /// final child. A plain append is fine for a document this library created, but every real
    /// document opened from disk carries a <c>w:sectPr</c>, and appending past it produces a file
    /// Word has to repair.
    /// </remarks>
    internal override void AppendBlock(OpenXmlElement element)
    {
        var sectionProperties = Element.GetFirstChild<WordLib.SectionProperties>();
        if (sectionProperties is null)
        {
            Element.AppendChild(element);
            return;
        }

        Element.InsertBefore(element, sectionProperties);
    }
}
