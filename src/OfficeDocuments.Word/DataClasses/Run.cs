using OfficeDocuments.Word.Formatting;
using OfficeDocuments.Word.Interfaces;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.DataClasses;

/// <summary>
/// A run: a span of text inside a paragraph that shares one set of character formatting.
/// </summary>
public class Run : IRun
{
    private readonly DocumentContext _context;

    internal WordLib.Run Element { get; }

    internal Run(WordLib.Run element, DocumentContext context)
    {
        ArgumentNullException.ThrowIfNull(element);

        Element = element;
        _context = context;
    }

    /// <inheritdoc />
    public string Text
    {
        get => RunContent.Read(Element);
        set
        {
            ArgumentNullException.ThrowIfNull(value);

            RunContent.Replace(Element, value);
        }
    }

    /// <inheritdoc />
    public TextFormat Format => RunFormatMapper.Read(Element);

    /// <inheritdoc />
    public IRun ApplyFormat(TextFormat format)
    {
        ArgumentNullException.ThrowIfNull(format);

        RunFormatMapper.Apply(Element, format, _context.EnsureStyle);

        return this;
    }
}
