using OfficeDocuments.Word.Formatting;

namespace OfficeDocuments.Word.Interfaces;

/// <summary>
/// A table row.
/// </summary>
public interface ITableRow
{
    /// <summary>
    /// The cells of this row, in document order.
    /// </summary>
    IReadOnlyList<ITableCell> Cells { get; }

    /// <summary>
    /// Appends a cell.
    /// </summary>
    /// <param name="text">Text for the cell's first paragraph, or <see langword="null"/> to leave it empty.</param>
    /// <param name="format">Cell formatting, or <see langword="null"/> for the table default.</param>
    /// <returns>The new cell.</returns>
    ITableCell AddCell(string? text = null, TableCellFormat? format = null);

    /// <summary>
    /// Marks this row as a header row, so Word repeats it at the top of every page the table spans.
    /// </summary>
    /// <param name="isHeader"><see langword="false"/> to clear the marking.</param>
    /// <returns>This row, for chaining.</returns>
    ITableRow RepeatAsHeader(bool isHeader = true);

    /// <summary>
    /// <see langword="true"/> when this row repeats as a header on each page.
    /// </summary>
    bool IsHeader { get; }

    /// <summary>
    /// The row's text, cells joined with <c>\t</c>.
    /// </summary>
    string GetAllTexts();
}
