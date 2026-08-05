using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;
using OfficeDocuments.Word.Interfaces;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.DataClasses;

/// <summary>
/// A paragraph: the block that holds runs of text and carries paragraph-level formatting.
/// </summary>
public class Paragraph : IParagraph
{
    private readonly DocumentContext _context;
    private readonly ElementWrapperList<WordLib.Run, IRun> _runs;

    internal WordLib.Paragraph Element { get; }

    /// <inheritdoc />
    public IReadOnlyList<IRun> Runs => _runs.Items;

    /// <inheritdoc />
    public ParagraphFormat Format => ParagraphFormatMapper.Read(Element, _context.ResolveListStyle);

    internal Paragraph(WordLib.Paragraph element, DocumentContext context)
    {
        ArgumentNullException.ThrowIfNull(element);

        Element = element;
        _context = context;

        // Descendants rather than direct children: a run inside a w:hyperlink is still one of this
        // paragraph's runs, and reading only direct children would hide the text of every link.
        _runs = new ElementWrapperList<WordLib.Run, IRun>(
            () => element.Descendants<WordLib.Run>(),
            run => new Run(run, context));
    }

    /// <inheritdoc />
    public IParagraph ApplyFormat(ParagraphFormat format)
    {
        ArgumentNullException.ThrowIfNull(format);

        ParagraphFormatMapper.Apply(Element, format, _context.EnsureStyle, _context.EnsureListNumbering);

        return this;
    }

    /// <inheritdoc />
    public IParagraph AddText(string text) => AddText(text, null);

    /// <inheritdoc />
    public IParagraph AddText(string text, TextFormat? format)
    {
        AddRun(text, format);

        return this;
    }

    /// <inheritdoc />
    public IRun AddRun(string text, TextFormat? format = null)
    {
        ArgumentNullException.ThrowIfNull(text);

        var element = new WordLib.Run();
        RunFormatMapper.Apply(element, format, _context.EnsureStyle);
        RunContent.Append(element, text);
        Element.AppendChild(element);

        return _runs.Wrap(element);
    }

    /// <inheritdoc />
    public IParagraph AddBreak(BreakType type)
    {
        var element = new WordLib.Run();
        element.AppendChild(new WordLib.Break { Type = ToOpenXml(type) });
        Element.AppendChild(element);
        _runs.Wrap(element);

        return this;
    }

    /// <inheritdoc />
    public IParagraph AddHyperlink(string text, string url, TextFormat? format = null)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentException.ThrowIfNullOrWhiteSpace(url);

        if (!Uri.TryCreate(url, UriKind.Absolute, out var target))
        {
            throw new ArgumentException($"'{url}' is not an absolute URI.", nameof(url));
        }

        var relationshipId = _context.CreateExternalRelationship(Element, target);

        // The hyperlink style is the base and the caller's format layers on top, so a caller can
        // recolour a link without losing the fact that it is one.
        var runFormat = new TextFormat { StyleId = WordStyleIds.Hyperlink }.Merge(format);

        var run = new WordLib.Run();
        RunFormatMapper.Apply(run, runFormat, _context.EnsureStyle);
        RunContent.Append(run, text);

        // The run sits inside the w:hyperlink container, which is why Runs reads descendants.
        Element.AppendChild(new WordLib.Hyperlink(run) { Id = relationshipId });
        _runs.Wrap(run);

        return this;
    }

    /// <inheritdoc />
    public IParagraph AddImage(Stream content, ImageSize? size = null, string? description = null)
        => AddImage(content, imageType: null, size, description, name: null);

    /// <inheritdoc />
    public IParagraph AddImage(Stream content, ImageType imageType, ImageSize? size = null, string? description = null)
        => AddImage(content, imageType, size, description, name: null);

    /// <inheritdoc />
    public IParagraph AddImage(string filePath, ImageSize? size = null, string? description = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        var imageType = InlineImageBuilder.InferType(filePath);

        using var content = File.OpenRead(filePath);

        return AddImage(content, imageType, size, description, Path.GetFileName(filePath));
    }

    private IParagraph AddImage(Stream content, ImageType? imageType, ImageSize? size, string? description, string? name)
    {
        ArgumentNullException.ThrowIfNull(content);

        var metadata = ImageMetadata.TryRead(content);
        var resolvedType = imageType
            ?? metadata?.Type
            ?? throw new ArgumentException(
                "Could not determine the image format. Pass the image type explicitly.",
                nameof(content));

        var (widthInPoints, heightInPoints) = ResolveSize(size ?? ImageSize.Intrinsic, metadata);

        // TryRead restores the position it found, which for a fresh stream is the start; rewinding
        // here makes the whole file reach the part even when the caller had already read from it.
        if (content.CanSeek)
        {
            content.Position = 0;
        }

        var relationshipId = _context.AddImagePart(Element, content, InlineImageBuilder.ToPartType(resolvedType));
        var drawingId = _context.NextDrawingId();

        var run = InlineImageBuilder.CreateRun(
            relationshipId,
            drawingId,
            name ?? $"Picture {drawingId}",
            description,
            widthInPoints,
            heightInPoints);

        Element.AppendChild(run);
        _runs.Wrap(run);

        return this;
    }

    /// <summary>
    /// Turns a requested size into concrete dimensions, using the image's own size where needed.
    /// </summary>
    private static (double Width, double Height) ResolveSize(ImageSize size, ImageMetadata? metadata)
    {
        if (size.WidthInPoints is { } width && size.HeightInPoints is { } height)
        {
            return (width, height);
        }

        if (metadata is null)
        {
            throw new ArgumentException(
                "The image's own size could not be read, so the size has to be given as ImageSize.Exact.",
                nameof(size));
        }

        // Deriving the missing dimension from the pixel ratio rather than from the points keeps the
        // result exact regardless of the file's stated resolution.
        if (size.WidthInPoints is { } widthOnly)
        {
            return (widthOnly, widthOnly * metadata.PixelHeight / metadata.PixelWidth);
        }

        if (size.HeightInPoints is { } heightOnly)
        {
            return (heightOnly * metadata.PixelWidth / metadata.PixelHeight, heightOnly);
        }

        return (metadata.WidthInPoints, metadata.HeightInPoints);
    }

    /// <inheritdoc />
    public IParagraph SetText(string text, TextFormat? format = null)
    {
        ArgumentNullException.ThrowIfNull(text);

        // Everything but w:pPr goes, so the paragraph keeps its style, alignment, and list membership
        // while losing its content. Removing the properties too would silently reset the formatting of
        // every paragraph a template fill touches.
        foreach (var child in Element.ChildElements.Where(child => child is not WordLib.ParagraphProperties).ToList())
        {
            child.Remove();
        }

        return AddText(text, format);
    }

    /// <inheritdoc />
    public int ReplaceText(string oldValue, string newValue, StringComparison comparison = StringComparison.Ordinal)
    {
        ArgumentException.ThrowIfNullOrEmpty(oldValue);
        ArgumentNullException.ThrowIfNull(newValue);

        return TextReplacer.Replace(Element, oldValue, newValue, comparison);
    }

    /// <inheritdoc />
    public IEnumerable<IText> GetTextElements()
    {
        foreach (var element in Element.Descendants<WordLib.Text>())
        {
            yield return new Text(element);
        }
    }

    /// <inheritdoc />
    public string GetTexts() => RunContent.Read(Element);

    private static WordLib.BreakValues ToOpenXml(BreakType type)
    {
        return type switch
        {
            BreakType.Page => WordLib.BreakValues.Page,
            BreakType.Column => WordLib.BreakValues.Column,
            BreakType.TextWrapping => WordLib.BreakValues.TextWrapping,
            _ => throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported break type."),
        };
    }
}
