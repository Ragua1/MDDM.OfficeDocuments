using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;

namespace OfficeDocuments.Word.Interfaces;

/// <summary>
/// Anything that holds block-level content: the document body, a header, a footer, or a table cell.
/// </summary>
/// <remarks>
/// WordprocessingML uses the same block content model in all of those places, so they share one
/// contract here rather than each growing its own paragraph and table API.
/// </remarks>
public interface IBlockContainer
{
    /// <summary>
    /// The paragraphs directly inside this container, in document order.
    /// </summary>
    /// <remarks>
    /// Projected from the document, so a paragraph added through this interface is visible here
    /// immediately. Paragraphs nested inside a table are reached through that table's cells, not here.
    /// </remarks>
    IReadOnlyList<IParagraph> Paragraphs { get; }

    /// <summary>
    /// The tables directly inside this container, in document order.
    /// </summary>
    IReadOnlyList<ITable> Tables { get; }

    /// <summary>
    /// Appends an empty paragraph.
    /// </summary>
    /// <returns>The new paragraph.</returns>
    IParagraph AddParagraph();

    /// <summary>
    /// Appends an empty paragraph with the given formatting.
    /// </summary>
    /// <param name="format">Paragraph formatting, or <see langword="null"/> for the document default.</param>
    /// <returns>The new paragraph.</returns>
    IParagraph AddParagraph(ParagraphFormat? format);

    /// <summary>
    /// Appends a paragraph containing <paramref name="text"/>.
    /// </summary>
    /// <param name="text">Text of the paragraph. Newlines become line breaks.</param>
    /// <returns>The new paragraph.</returns>
    IParagraph AddParagraph(string text);

    /// <summary>
    /// Appends a formatted paragraph containing <paramref name="text"/>.
    /// </summary>
    /// <param name="text">Text of the paragraph. Newlines become line breaks.</param>
    /// <param name="format">Paragraph formatting, or <see langword="null"/> for the document default.</param>
    /// <param name="textFormat">Character formatting for the text, or <see langword="null"/> to inherit.</param>
    /// <returns>The new paragraph.</returns>
    IParagraph AddParagraph(string text, ParagraphFormat? format, TextFormat? textFormat = null);

    /// <summary>
    /// Appends a heading, defining the matching built-in style if the document lacks it.
    /// </summary>
    /// <param name="text">Heading text.</param>
    /// <param name="level">Heading level, 1 to 6.</param>
    /// <returns>The new paragraph.</returns>
    /// <exception cref="ArgumentOutOfRangeException">The level is outside 1 to 6.</exception>
    IParagraph AddHeading(string text, int level);

    /// <summary>
    /// Appends a list item, defining the numbering the list needs if the document lacks it.
    /// </summary>
    /// <param name="text">Item text.</param>
    /// <param name="style">Bullet or numbered.</param>
    /// <param name="level">Nesting depth, 0 for the outermost level.</param>
    /// <returns>The new paragraph.</returns>
    /// <exception cref="ArgumentOutOfRangeException">The level is outside 0 to 8.</exception>
    IParagraph AddListItem(string text, ListStyle style = ListStyle.Bullet, int level = 0);

    /// <summary>
    /// Appends an empty table of the given size, with one empty paragraph in each cell.
    /// </summary>
    /// <param name="rowCount">Number of rows, at least 1.</param>
    /// <param name="columnCount">Number of columns, at least 1.</param>
    /// <param name="format">Table formatting, or <see langword="null"/> for the document default.</param>
    /// <returns>The new table.</returns>
    /// <exception cref="ArgumentOutOfRangeException">A count is less than 1.</exception>
    ITable AddTable(int rowCount, int columnCount, TableFormat? format = null);

    /// <summary>
    /// Appends a table filled from <paramref name="rows"/>, sized to the longest row.
    /// </summary>
    /// <param name="rows">Cell text, row by row.</param>
    /// <param name="format">Table formatting, or <see langword="null"/> for the document default.</param>
    /// <returns>The new table.</returns>
    /// <exception cref="ArgumentException"><paramref name="rows"/> is empty or has no columns.</exception>
    ITable AddTable(IEnumerable<IEnumerable<string>> rows, TableFormat? format = null);

    /// <summary>
    /// Every paragraph in this container in document order, including those inside tables.
    /// </summary>
    /// <remarks>
    /// The navigation counterpart to <see cref="Paragraphs"/>, which stops at this container's own
    /// children. Table content is reached by descending through the cells, at any nesting depth, so a
    /// document-wide pass over the text does not have to know the table structure to visit all of it.
    /// </remarks>
    IEnumerable<IParagraph> GetAllParagraphs();

    /// <summary>
    /// The paragraphs whose text contains <paramref name="text"/>, including those inside tables.
    /// </summary>
    /// <param name="text">Text to look for.</param>
    /// <param name="comparison">How to compare. Ordinal by default.</param>
    /// <returns>The matching paragraphs, in document order.</returns>
    IEnumerable<IParagraph> FindParagraphs(string text, StringComparison comparison = StringComparison.Ordinal);

    /// <summary>
    /// Replaces every occurrence of <paramref name="oldValue"/> throughout this container, tables
    /// included.
    /// </summary>
    /// <remarks>
    /// Each paragraph is treated as one text, so a match is found even where Word split it across runs,
    /// and no match spans two paragraphs. See <see cref="IParagraph.ReplaceText"/> for what that means.
    /// </remarks>
    /// <param name="oldValue">Text to find.</param>
    /// <param name="newValue">Replacement text. Newlines become line breaks; empty text deletes.</param>
    /// <param name="comparison">How to compare. Ordinal by default.</param>
    /// <returns>The number of occurrences replaced.</returns>
    /// <exception cref="ArgumentException"><paramref name="oldValue"/> is empty.</exception>
    int ReplaceText(string oldValue, string newValue, StringComparison comparison = StringComparison.Ordinal);

    /// <summary>
    /// Removes <paramref name="paragraph"/> from this container.
    /// </summary>
    /// <param name="paragraph">Paragraph to remove.</param>
    /// <returns>
    /// <see langword="true"/> if it was removed, <see langword="false"/> if it is not a direct child of
    /// this container.
    /// </returns>
    bool Remove(IParagraph paragraph);

    /// <summary>
    /// Removes <paramref name="table"/> from this container.
    /// </summary>
    /// <param name="table">Table to remove.</param>
    /// <returns>
    /// <see langword="true"/> if it was removed, <see langword="false"/> if it is not a direct child of
    /// this container.
    /// </returns>
    bool Remove(ITable table);

    /// <summary>
    /// The text of this container's block content, joined with <c>\n</c>.
    /// </summary>
    /// <remarks>
    /// Paragraphs and tables are read in document order. Table cells within a row are separated by
    /// <c>\t</c> and rows by <c>\n</c>, so the result keeps the shape of the content.
    /// </remarks>
    string GetAllTexts();
}
