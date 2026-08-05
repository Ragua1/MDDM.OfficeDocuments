using OfficeDocuments.Word.Formatting;

namespace OfficeDocuments.Word.Interfaces;

/// <summary>
/// A table: rows of cells, each of which is itself a block container.
/// </summary>
public interface ITable
{
    /// <summary>
    /// The rows of this table, in document order.
    /// </summary>
    IReadOnlyList<ITableRow> Rows { get; }

    /// <summary>
    /// Number of grid columns the table declares.
    /// </summary>
    /// <remarks>
    /// The declared grid, not the cell count of any particular row: a row containing a cell that spans
    /// two columns holds fewer cells than this.
    /// </remarks>
    int ColumnCount { get; }

    /// <summary>
    /// The table formatting this library models.
    /// </summary>
    TableFormat Format { get; }

    /// <summary>
    /// Applies the properties <paramref name="format"/> sets, leaving the others as they are.
    /// </summary>
    /// <param name="format">Formatting to apply.</param>
    /// <returns>This table, for chaining.</returns>
    ITable ApplyFormat(TableFormat format);

    /// <summary>
    /// Appends a row with one empty cell per grid column.
    /// </summary>
    /// <returns>The new row.</returns>
    ITableRow AddRow();

    /// <summary>
    /// Appends a row filled with <paramref name="cells"/>, padded with empty cells to the grid width.
    /// </summary>
    /// <param name="cells">Cell text, left to right.</param>
    /// <returns>The new row.</returns>
    ITableRow AddRow(params string[] cells);

    /// <summary>
    /// Removes <paramref name="row"/> from this table.
    /// </summary>
    /// <remarks>
    /// The grid is left alone, because it describes the table's columns rather than the removed row.
    /// Removing every row leaves a table Word renders as nothing; remove the table itself instead.
    /// </remarks>
    /// <param name="row">Row to remove.</param>
    /// <returns>
    /// <see langword="true"/> if it was removed, <see langword="false"/> if it is not a row of this
    /// table.
    /// </returns>
    bool Remove(ITableRow row);

    /// <summary>
    /// The cell at the given position.
    /// </summary>
    /// <param name="rowIndex">Zero-based row index.</param>
    /// <param name="columnIndex">Zero-based index within that row's cells.</param>
    /// <exception cref="ArgumentOutOfRangeException">Either index is outside the table.</exception>
    ITableCell GetCell(int rowIndex, int columnIndex);

    /// <summary>
    /// The table's text: cells joined with <c>\t</c>, rows joined with <c>\n</c>.
    /// </summary>
    string GetAllTexts();
}
