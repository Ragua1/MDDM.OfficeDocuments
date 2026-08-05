using System.Text;
using OfficeDocuments.Word.Formatting;
using OfficeDocuments.Word.Interfaces;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.DataClasses;

/// <summary>
/// A table row.
/// </summary>
public class TableRow : ITableRow
{
    private readonly ElementWrapperList<WordLib.TableCell, ITableCell> _cells;

    internal WordLib.TableRow Element { get; }

    /// <inheritdoc />
    public IReadOnlyList<ITableCell> Cells => _cells.Items;

    /// <inheritdoc />
    public bool IsHeader
    {
        get
        {
            var header = Element.GetFirstChild<WordLib.TableRowProperties>()
                ?.GetFirstChild<WordLib.TableHeader>();

            if (header is null)
            {
                return false;
            }

            // w:tblHeader takes the on/off vocabulary rather than a boolean, and a present element with
            // no value means "on".
            var value = header.Val;

            return value is null || value.Value == WordLib.OnOffOnlyValues.On;
        }
    }

    internal TableRow(WordLib.TableRow element, DocumentContext context)
    {
        ArgumentNullException.ThrowIfNull(element);

        Element = element;
        _cells = new ElementWrapperList<WordLib.TableCell, ITableCell>(
            () => element.Elements<WordLib.TableCell>(),
            cell => new TableCell(cell, context));
    }

    /// <inheritdoc />
    public ITableCell AddCell(string? text = null, TableCellFormat? format = null)
    {
        var element = TableCell.CreateElement();
        Element.AppendChild(element);

        var cell = (TableCell)_cells.Wrap(element);

        if (format is not null)
        {
            cell.ApplyFormat(format);
        }

        if (text is not null)
        {
            cell.SetText(text);
        }

        return cell;
    }

    /// <inheritdoc />
    public ITableRow RepeatAsHeader(bool isHeader = true)
    {
        var properties = GetOrCreateProperties();
        var existing = properties.GetFirstChild<WordLib.TableHeader>();
        existing?.Remove();

        properties.AppendChild(isHeader
            ? new WordLib.TableHeader()
            : new WordLib.TableHeader { Val = WordLib.OnOffOnlyValues.Off });

        return this;
    }

    /// <inheritdoc />
    public string GetAllTexts()
    {
        var builder = new StringBuilder();
        var isFirstCell = true;

        foreach (var cell in Element.Elements<WordLib.TableCell>())
        {
            if (!isFirstCell)
            {
                builder.Append('\t');
            }

            builder.Append(RunContent.Read(cell));
            isFirstCell = false;
        }

        return builder.ToString();
    }

    /// <summary>
    /// Returns the row's properties element, creating it in its required first position.
    /// </summary>
    private WordLib.TableRowProperties GetOrCreateProperties()
    {
        var existing = Element.GetFirstChild<WordLib.TableRowProperties>();
        if (existing is not null)
        {
            return existing;
        }

        var properties = new WordLib.TableRowProperties();
        Element.InsertAt(properties, 0);

        return properties;
    }
}
