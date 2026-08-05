using System.Text;
using OfficeDocuments.Word.Formatting;
using OfficeDocuments.Word.Interfaces;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.DataClasses;

/// <summary>
/// A table: rows of cells, each of which is itself a block container.
/// </summary>
public class Table : ITable
{
    private readonly DocumentContext _context;
    private readonly ElementWrapperList<WordLib.TableRow, ITableRow> _rows;

    internal WordLib.Table Element { get; }

    /// <inheritdoc />
    public IReadOnlyList<ITableRow> Rows => _rows.Items;

    /// <inheritdoc />
    public int ColumnCount => Element.GetFirstChild<WordLib.TableGrid>()?.Elements<WordLib.GridColumn>().Count() ?? 0;

    /// <inheritdoc />
    public TableFormat Format => TableFormatMapper.Read(Element);

    internal Table(WordLib.Table element, DocumentContext context)
    {
        ArgumentNullException.ThrowIfNull(element);

        Element = element;
        _context = context;
        _rows = new ElementWrapperList<WordLib.TableRow, ITableRow>(
            () => element.Elements<WordLib.TableRow>(),
            row => new TableRow(row, context));
    }

    /// <summary>
    /// Builds a table element with the grid its columns need.
    /// </summary>
    /// <remarks>
    /// <c>CT_Tbl</c> requires a <c>w:tblGrid</c>, and it must declare one <c>w:gridCol</c> per column
    /// or Word cannot lay the table out. Creating it up front is what makes the table well-formed
    /// before any row exists.
    /// </remarks>
    internal static WordLib.Table CreateElement(int columnCount)
    {
        var element = new WordLib.Table();

        // A default single-border look, because a table with no borders at all is rarely what a caller
        // producing a business document wants, and it is overridable through TableFormat.
        TableFormatMapper.Apply(element, new TableFormat { Borders = Enums.TableBorders.All, WidthPercent = 100 });

        var grid = new WordLib.TableGrid();
        for (var column = 0; column < columnCount; column++)
        {
            grid.AppendChild(new WordLib.GridColumn());
        }

        element.AppendChild(grid);

        return element;
    }

    /// <summary>
    /// Reads a table's text without needing a wrapper: cells joined with tabs, rows with newlines.
    /// </summary>
    internal static string ReadText(WordLib.Table element)
    {
        var builder = new StringBuilder();
        var isFirstRow = true;

        foreach (var row in element.Elements<WordLib.TableRow>())
        {
            if (!isFirstRow)
            {
                builder.Append('\n');
            }

            var isFirstCell = true;
            foreach (var cell in row.Elements<WordLib.TableCell>())
            {
                if (!isFirstCell)
                {
                    builder.Append('\t');
                }

                builder.Append(RunContent.Read(cell));
                isFirstCell = false;
            }

            isFirstRow = false;
        }

        return builder.ToString();
    }

    /// <inheritdoc />
    public ITable ApplyFormat(TableFormat format)
    {
        ArgumentNullException.ThrowIfNull(format);

        TableFormatMapper.Apply(Element, format);

        return this;
    }

    /// <inheritdoc />
    public ITableRow AddRow() => AddRow([]);

    /// <inheritdoc />
    public ITableRow AddRow(params string[] cells)
    {
        ArgumentNullException.ThrowIfNull(cells);

        var element = new WordLib.TableRow();
        Element.AppendChild(element);

        var row = (TableRow)_rows.Wrap(element);

        // Padded to the grid width: a row with fewer cells than the grid renders as a ragged table.
        var cellCount = Math.Max(cells.Length, ColumnCount);
        for (var index = 0; index < cellCount; index++)
        {
            row.AddCell(index < cells.Length ? cells[index] : null);
        }

        return row;
    }

    /// <inheritdoc />
    public bool Remove(ITableRow row)
    {
        ArgumentNullException.ThrowIfNull(row);

        if (row is not TableRow implementation || !ReferenceEquals(implementation.Element.Parent, Element))
        {
            return false;
        }

        implementation.Element.Remove();

        return true;
    }

    /// <summary>
    /// Every paragraph in this table, cell by cell, including those in nested tables.
    /// </summary>
    /// <remarks>
    /// Internal because a cell is already an <see cref="IBlockContainer"/>, so a caller who wants a
    /// table's paragraphs can reach them through the cells. This exists for the container-level walk,
    /// which has to descend into a table without knowing it is one.
    /// </remarks>
    internal IEnumerable<IParagraph> GetAllParagraphs()
    {
        foreach (var row in Rows)
        {
            foreach (var cell in row.Cells)
            {
                foreach (var paragraph in cell.GetAllParagraphs())
                {
                    yield return paragraph;
                }
            }
        }
    }

    /// <inheritdoc />
    public ITableCell GetCell(int rowIndex, int columnIndex)
    {
        var rows = Rows;
        if (rowIndex < 0 || rowIndex >= rows.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(rowIndex), rowIndex, $"The table has {rows.Count} rows.");
        }

        var cells = rows[rowIndex].Cells;
        if (columnIndex < 0 || columnIndex >= cells.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(columnIndex), columnIndex, $"Row {rowIndex} has {cells.Count} cells.");
        }

        return cells[columnIndex];
    }

    /// <inheritdoc />
    public string GetAllTexts() => ReadText(Element);
}
