using OfficeDocuments.Word.Formatting;
using OfficeDocuments.Word.Interfaces;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.DataClasses;

/// <summary>
/// A table cell. Holds block content, so it takes paragraphs, headings, lists, and nested tables.
/// </summary>
public class TableCell : BlockContainer, ITableCell
{
    internal new WordLib.TableCell Element { get; }

    /// <inheritdoc />
    public TableCellFormat Format => TableCellFormatMapper.Read(Element);

    internal TableCell(WordLib.TableCell element, DocumentContext context) : base(element, context)
    {
        Element = element;
    }

    /// <summary>
    /// Builds a cell element that already satisfies the schema.
    /// </summary>
    /// <remarks>
    /// <c>CT_Tc</c> requires block content, so a cell with no paragraph is invalid and Word offers to
    /// repair the document. Creating the paragraph with the cell means a cell is never in that state.
    /// </remarks>
    internal static WordLib.TableCell CreateElement()
    {
        var element = new WordLib.TableCell();
        element.AppendChild(new WordLib.Paragraph());

        return element;
    }

    /// <inheritdoc />
    public ITableCell ApplyFormat(TableCellFormat format)
    {
        ArgumentNullException.ThrowIfNull(format);

        TableCellFormatMapper.Apply(Element, format);

        return this;
    }

    /// <inheritdoc />
    public ITableCell SetText(string text, TextFormat? format = null)
    {
        ArgumentNullException.ThrowIfNull(text);

        foreach (var paragraph in Element.Elements<WordLib.Paragraph>().ToList())
        {
            paragraph.Remove();
        }

        foreach (var table in Element.Elements<WordLib.Table>().ToList())
        {
            table.Remove();
        }

        AddParagraph(text, format: null, textFormat: format);

        return this;
    }
}
