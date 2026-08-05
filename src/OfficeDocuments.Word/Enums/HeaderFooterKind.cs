namespace OfficeDocuments.Word.Enums;

/// <summary>
/// Which pages a header or footer applies to.
/// </summary>
/// <remarks>
/// <see cref="First"/> and <see cref="Even"/> each need a document-level switch turned on before Word
/// shows them; the library sets that switch when one is added, so choosing a kind here is enough.
/// </remarks>
public enum HeaderFooterKind
{
    /// <summary>
    /// Every page that no more specific header or footer covers.
    /// </summary>
    Default,
    /// <summary>
    /// The first page only, for a title page or letterhead.
    /// </summary>
    First,
    /// <summary>
    /// Even-numbered pages, for documents printed as facing pages.
    /// </summary>
    Even,
}
