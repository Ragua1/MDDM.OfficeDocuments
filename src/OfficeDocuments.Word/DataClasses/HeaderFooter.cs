using DocumentFormat.OpenXml;
using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Interfaces;

namespace OfficeDocuments.Word.DataClasses;

/// <summary>
/// A header or footer. Holds block content, so it takes paragraphs, tables, and images.
/// </summary>
public class HeaderFooter : BlockContainer, IHeaderFooter
{
    internal HeaderFooter(OpenXmlCompositeElement element, DocumentContext context, HeaderFooterKind kind, bool isHeader)
        : base(element, context)
    {
        Kind = kind;
        IsHeader = isHeader;
    }

    /// <inheritdoc />
    public HeaderFooterKind Kind { get; }

    /// <inheritdoc />
    public bool IsHeader { get; }
}
