using System.Text;
using DocumentFormat.OpenXml;
using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;
using OfficeDocuments.Word.Interfaces;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.DataClasses;

/// <summary>
/// Shared implementation of block-level content for the body, headers, footers, and table cells.
/// </summary>
/// <remarks>
/// These four containers use the same block content model in WordprocessingML, so they share it here.
/// The only behaviour that genuinely differs is where a new block goes, which is why
/// <see cref="AppendBlock"/> is the single overridable point.
/// </remarks>
public abstract class BlockContainer : IBlockContainer
{
    private readonly ElementWrapperList<WordLib.Paragraph, IParagraph> _paragraphs;
    private readonly ElementWrapperList<WordLib.Table, ITable> _tables;

    internal OpenXmlCompositeElement Element { get; }

    internal DocumentContext Context { get; }

    /// <inheritdoc />
    public IReadOnlyList<IParagraph> Paragraphs => _paragraphs.Items;

    /// <inheritdoc />
    public IReadOnlyList<ITable> Tables => _tables.Items;

    internal BlockContainer(OpenXmlCompositeElement element, DocumentContext context)
    {
        ArgumentNullException.ThrowIfNull(element);

        Element = element;
        Context = context;
        _paragraphs = new ElementWrapperList<WordLib.Paragraph, IParagraph>(
            () => element.Elements<WordLib.Paragraph>(),
            paragraph => new Paragraph(paragraph, context));
        _tables = new ElementWrapperList<WordLib.Table, ITable>(
            () => element.Elements<WordLib.Table>(),
            table => new Table(table, context));
    }

    /// <inheritdoc />
    public IParagraph AddParagraph() => AddParagraph(format: null);

    /// <inheritdoc />
    public IParagraph AddParagraph(ParagraphFormat? format)
    {
        var element = new WordLib.Paragraph();
        AppendBlock(element);

        var paragraph = _paragraphs.Wrap(element);

        return format is null ? paragraph : paragraph.ApplyFormat(format);
    }

    /// <inheritdoc />
    public IParagraph AddParagraph(string text) => AddParagraph(text, format: null);

    /// <inheritdoc />
    public IParagraph AddParagraph(string text, ParagraphFormat? format, TextFormat? textFormat = null)
    {
        ArgumentNullException.ThrowIfNull(text);

        return AddParagraph(format).AddText(text, textFormat);
    }

    /// <inheritdoc />
    public IParagraph AddHeading(string text, int level)
    {
        ArgumentNullException.ThrowIfNull(text);

        return AddParagraph(text, new ParagraphFormat { StyleId = WordStyleIds.Heading(level) });
    }

    /// <inheritdoc />
    public IParagraph AddListItem(string text, ListStyle style = ListStyle.Bullet, int level = 0)
    {
        ArgumentNullException.ThrowIfNull(text);

        return AddParagraph(text, new ParagraphFormat { ListStyle = style, ListLevel = level });
    }

    /// <inheritdoc />
    public ITable AddTable(int rowCount, int columnCount, TableFormat? format = null)
    {
        if (rowCount < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(rowCount), rowCount, "A table needs at least one row.");
        }

        if (columnCount < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(columnCount), columnCount, "A table needs at least one column.");
        }

        var table = CreateTable(columnCount, format);
        for (var row = 0; row < rowCount; row++)
        {
            table.AddRow();
        }

        return table;
    }

    /// <inheritdoc />
    public ITable AddTable(IEnumerable<IEnumerable<string>> rows, TableFormat? format = null)
    {
        ArgumentNullException.ThrowIfNull(rows);

        var materializedRows = rows.Select(row => row?.ToArray() ?? []).ToArray();
        if (materializedRows.Length == 0)
        {
            throw new ArgumentException("A table needs at least one row.", nameof(rows));
        }

        // The grid has to be wide enough for the longest row, or the extra cells fall outside it.
        var columnCount = materializedRows.Max(row => row.Length);
        if (columnCount == 0)
        {
            throw new ArgumentException("A table needs at least one column.", nameof(rows));
        }

        var table = CreateTable(columnCount, format);
        foreach (var row in materializedRows)
        {
            table.AddRow(row);
        }

        return table;
    }

    /// <inheritdoc />
    public IEnumerable<IParagraph> GetAllParagraphs()
    {
        // Walked over the children rather than over Paragraphs and Tables separately, so the order is
        // the document's own and a paragraph between two tables is not reported out of place.
        foreach (var child in Element.ChildElements)
        {
            switch (child)
            {
                case WordLib.Paragraph paragraph:
                    yield return _paragraphs.Wrap(paragraph);
                    break;

                case WordLib.Table table:
                    foreach (var nested in ((Table)_tables.Wrap(table)).GetAllParagraphs())
                    {
                        yield return nested;
                    }

                    break;
            }
        }
    }

    /// <inheritdoc />
    public IEnumerable<IParagraph> FindParagraphs(string text, StringComparison comparison = StringComparison.Ordinal)
    {
        ArgumentNullException.ThrowIfNull(text);

        return GetAllParagraphs().Where(paragraph => paragraph.GetTexts().Contains(text, comparison));
    }

    /// <inheritdoc />
    public int ReplaceText(string oldValue, string newValue, StringComparison comparison = StringComparison.Ordinal)
    {
        ArgumentException.ThrowIfNullOrEmpty(oldValue);
        ArgumentNullException.ThrowIfNull(newValue);

        // Materialized first: replacing removes and inserts elements, and GetAllParagraphs reads the
        // tree lazily, so enumerating it while editing would walk a collection that is moving.
        var paragraphs = GetAllParagraphs().ToList();

        return paragraphs.Sum(paragraph => paragraph.ReplaceText(oldValue, newValue, comparison));
    }

    /// <inheritdoc />
    public bool Remove(IParagraph paragraph)
    {
        ArgumentNullException.ThrowIfNull(paragraph);

        return RemoveChild(paragraph is Paragraph implementation ? implementation.Element : null);
    }

    /// <inheritdoc />
    public bool Remove(ITable table)
    {
        ArgumentNullException.ThrowIfNull(table);

        return RemoveChild(table is Table implementation ? implementation.Element : null);
    }

    /// <summary>
    /// Removes <paramref name="element"/> if it really is a child of this container.
    /// </summary>
    /// <remarks>
    /// The ownership check is the point. <c>Remove</c> on a paragraph belonging to another container
    /// would otherwise take it out of that one, which is a silent edit to a part of the document the
    /// caller did not name.
    /// </remarks>
    private bool RemoveChild(OpenXmlElement? element)
    {
        if (element is null || !ReferenceEquals(element.Parent, Element))
        {
            return false;
        }

        element.Remove();

        return true;
    }

    /// <inheritdoc />
    public string GetAllTexts()
    {
        var builder = new StringBuilder();
        var isFirstBlock = true;

        // Walked in document order rather than paragraphs-then-tables, so the result reads the way the
        // document does.
        foreach (var child in Element.ChildElements)
        {
            var blockText = child switch
            {
                WordLib.Paragraph paragraph => RunContent.Read(paragraph),
                WordLib.Table table => Table.ReadText(table),
                _ => null,
            };

            if (blockText is null)
            {
                continue;
            }

            if (!isFirstBlock)
            {
                builder.Append('\n');
            }

            builder.Append(blockText);
            isFirstBlock = false;
        }

        return builder.ToString();
    }

    /// <summary>
    /// Places a new block-level element inside this container.
    /// </summary>
    /// <remarks>
    /// Overridden where the container has a schema constraint on the tail of its content, as the body
    /// does with <c>w:sectPr</c>.
    /// </remarks>
    internal virtual void AppendBlock(OpenXmlElement element) => Element.AppendChild(element);

    private Table CreateTable(int columnCount, TableFormat? format)
    {
        var element = Table.CreateElement(columnCount);
        AppendBlock(element);

        var table = (Table)_tables.Wrap(element);
        if (format is not null)
        {
            table.ApplyFormat(format);
        }

        return table;
    }
}
