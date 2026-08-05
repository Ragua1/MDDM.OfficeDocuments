using OfficeDocuments.Word.Formatting;

namespace OfficeDocuments.Word.Interfaces;

/// <summary>
/// A table cell. Holds block content, so it takes paragraphs, headings, lists, and nested tables.
/// </summary>
public interface ITableCell : IBlockContainer
{
    /// <summary>
    /// The cell formatting this library models.
    /// </summary>
    TableCellFormat Format { get; }

    /// <summary>
    /// Applies the properties <paramref name="format"/> sets, leaving the others as they are.
    /// </summary>
    /// <param name="format">Formatting to apply.</param>
    /// <returns>This cell, for chaining.</returns>
    ITableCell ApplyFormat(TableCellFormat format);

    /// <summary>
    /// Replaces the cell's content with a single paragraph containing <paramref name="text"/>.
    /// </summary>
    /// <remarks>
    /// A cell always keeps at least one paragraph, because WordprocessingML requires block content in
    /// every cell and Word repairs a document that has an empty one.
    /// </remarks>
    /// <param name="text">Text to write. Newlines become line breaks.</param>
    /// <param name="format">Character formatting, or <see langword="null"/> to inherit.</param>
    /// <returns>This cell, for chaining.</returns>
    ITableCell SetText(string text, TextFormat? format = null);
}
