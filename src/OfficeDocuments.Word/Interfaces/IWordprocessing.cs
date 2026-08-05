using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;

namespace OfficeDocuments.Word.Interfaces;

/// <summary>
/// A Word document: its content, its page setup, and its metadata.
/// </summary>
public interface IWordprocessing : IDisposable
{
    /// <summary>
    /// The document body.
    /// </summary>
    /// <returns>The same instance on every call.</returns>
    IBody GetBody();

    /// <summary>
    /// Adds a header, or returns the existing one for <paramref name="kind"/>.
    /// </summary>
    /// <remarks>
    /// A <see cref="HeaderFooterKind.First"/> or <see cref="HeaderFooterKind.Even"/> header needs a
    /// document-level switch turned on before Word displays it. That is done here, so the returned
    /// header renders without further setup.
    /// </remarks>
    /// <param name="kind">Which pages the header applies to.</param>
    /// <returns>The header's content container.</returns>
    IHeaderFooter AddHeader(HeaderFooterKind kind = HeaderFooterKind.Default);

    /// <summary>
    /// Adds a footer, or returns the existing one for <paramref name="kind"/>.
    /// </summary>
    /// <param name="kind">Which pages the footer applies to.</param>
    /// <returns>The footer's content container.</returns>
    IHeaderFooter AddFooter(HeaderFooterKind kind = HeaderFooterKind.Default);

    /// <summary>
    /// The headers and footers this document defines, in the order the section references them.
    /// </summary>
    /// <remarks>
    /// Read from the document, so a document opened from disk reports the headers and footers it already
    /// had, not only those added through <see cref="AddHeader"/> and <see cref="AddFooter"/>. A header
    /// part the section does not reference is left out: nothing displays it.
    /// </remarks>
    IReadOnlyList<IHeaderFooter> HeadersAndFooters { get; }

    /// <summary>
    /// Replaces every occurrence of <paramref name="oldValue"/> in the body, in every table, and in
    /// every header and footer.
    /// </summary>
    /// <remarks>
    /// The template-filling entry point: a document's placeholders are rarely all in the body, and a
    /// date or customer name in a running header is exactly the kind of thing a body-only pass leaves
    /// behind. Matching works per paragraph, so it survives Word having split a placeholder across runs.
    /// </remarks>
    /// <param name="oldValue">Text to find.</param>
    /// <param name="newValue">Replacement text. Newlines become line breaks; empty text deletes.</param>
    /// <param name="comparison">How to compare. Ordinal by default.</param>
    /// <returns>The number of occurrences replaced.</returns>
    /// <exception cref="ArgumentException"><paramref name="oldValue"/> is empty.</exception>
    int ReplaceText(string oldValue, string newValue, StringComparison comparison = StringComparison.Ordinal);

    /// <summary>
    /// The page size, orientation, and margins the library models.
    /// </summary>
    PageSetup PageSetup { get; }

    /// <summary>
    /// Applies the page properties <paramref name="setup"/> sets, leaving the others as they are.
    /// </summary>
    /// <param name="setup">Page setup to apply.</param>
    /// <returns>This document, for chaining.</returns>
    IWordprocessing ApplyPageSetup(PageSetup setup);

    /// <summary>
    /// The document's core properties.
    /// </summary>
    DocumentMetadata Metadata { get; }

    /// <summary>
    /// Applies the properties <paramref name="metadata"/> sets, leaving the others as they are.
    /// </summary>
    /// <remarks>
    /// Assign an empty string to clear a property that is already set; <see langword="null"/> leaves it
    /// alone.
    /// </remarks>
    /// <param name="metadata">Metadata to apply.</param>
    /// <returns>This document, for chaining.</returns>
    IWordprocessing SetMetadata(DocumentMetadata metadata);

    /// <summary>
    /// Saves and closes the document. Safe to call more than once.
    /// </summary>
    /// <param name="saveDocument">
    /// <see langword="false"/> to discard changes. Ignored for a document opened read-only.
    /// </param>
    void Close(bool saveDocument = true);
}
