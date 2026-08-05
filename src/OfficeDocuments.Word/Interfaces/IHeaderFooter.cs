using OfficeDocuments.Word.Enums;

namespace OfficeDocuments.Word.Interfaces;

/// <summary>
/// A header or footer. Holds block content, so it takes paragraphs, tables, and images.
/// </summary>
public interface IHeaderFooter : IBlockContainer
{
    /// <summary>
    /// Which pages this header or footer applies to.
    /// </summary>
    HeaderFooterKind Kind { get; }

    /// <summary>
    /// <see langword="true"/> for a header, <see langword="false"/> for a footer.
    /// </summary>
    bool IsHeader { get; }
}
