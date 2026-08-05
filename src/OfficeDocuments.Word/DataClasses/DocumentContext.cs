using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.DataClasses;

/// <summary>
/// Document-level services that the body, paragraph, and run wrappers need but must not own.
/// </summary>
/// <remarks>
/// The wrappers below <see cref="Wordprocessing"/> are element-scoped: a <see cref="Paragraph"/>
/// knows its own <c>w:p</c> and nothing else. Several Word features are not element-scoped, though —
/// named styles and list numbering live in shared parts, hyperlinks need a relationship on the
/// document part, and images need a media part. Passing this context down the tree gives those
/// features one shared seam instead of forcing a second, parallel object model.
/// </remarks>
internal sealed class DocumentContext(MainDocumentPart mainDocumentPart)
{
    private readonly HashSet<string> _resolvedStyleIds = new(StringComparer.Ordinal);
    private readonly Dictionary<ListStyle, int> _listNumberingIds = [];
    private uint _nextDrawingId;
    private bool _drawingIdsSeeded;

    /// <summary>
    /// The main document part that owns every element reachable from this context.
    /// </summary>
    internal MainDocumentPart MainDocumentPart { get; } = mainDocumentPart;

    /// <summary>
    /// Makes sure a paragraph style referenced by <paramref name="styleId"/> is defined in the
    /// document, adding the built-in definition the first time it is used.
    /// </summary>
    /// <remarks>
    /// An identifier this library does not know is left alone rather than rejected: a document opened
    /// from a template legitimately references styles defined in that template, and inventing a
    /// definition for one would overwrite the template's look.
    /// </remarks>
    internal void EnsureStyle(string styleId)
    {
        if (!_resolvedStyleIds.Add(styleId) || !BuiltInStyles.IsKnown(styleId))
        {
            return;
        }

        var styles = GetOrCreateStyles();
        if (ContainsStyle(styles, styleId))
        {
            return;
        }

        // Every other built-in style is based on Normal, so it has to exist first.
        if (!string.Equals(styleId, WordStyleIds.Normal, StringComparison.Ordinal))
        {
            EnsureStyle(WordStyleIds.Normal);
        }

        var definition = BuiltInStyles.TryCreate(styleId);
        if (definition is not null)
        {
            styles.AppendChild(definition);
        }
    }

    /// <summary>
    /// Returns the numbering identifier for <paramref name="style"/>, defining the numbering in the
    /// document the first time that style is used.
    /// </summary>
    /// <returns>
    /// The numbering identifier to reference from a paragraph, or
    /// <see cref="ListNumbering.NoNumberingId"/> for <see cref="ListStyle.None"/>.
    /// </returns>
    internal int EnsureListNumbering(ListStyle style)
    {
        if (style == ListStyle.None)
        {
            return ListNumbering.NoNumberingId;
        }

        if (_listNumberingIds.TryGetValue(style, out var cached))
        {
            return cached;
        }

        var numbering = GetOrCreateNumbering();

        // An opened document may already carry a suitable definition. Reusing it keeps the numbering
        // part from growing a near-duplicate every time a document is appended to.
        if (FindExistingNumbering(numbering, style) is { } existing)
        {
            _listNumberingIds[style] = existing;
            return existing;
        }

        var abstractNumberingId = NextAbstractNumberingId(numbering);
        InsertAbstractNumbering(numbering, ListNumbering.CreateAbstractNumbering(style, abstractNumberingId));

        var numberingId = NextNumberingId(numbering);
        numbering.AppendChild(new WordLib.NumberingInstance
        {
            NumberID = numberingId,
            AbstractNumId = new WordLib.AbstractNumId { Val = abstractNumberingId },
        });

        _listNumberingIds[style] = numberingId;

        return numberingId;
    }

    /// <summary>
    /// Classifies a numbering identifier found in a document as a bullet or numbered list.
    /// </summary>
    /// <remarks>
    /// Resolved by inspecting the numbering definition rather than by remembering what this library
    /// wrote, so a list in a document opened from disk is classified too.
    /// </remarks>
    internal ListStyle? ResolveListStyle(int numberingId)
    {
        var numbering = MainDocumentPart.NumberingDefinitionsPart?.Numbering;
        if (numbering is null)
        {
            return null;
        }

        var abstractNumbering = FindAbstractNumbering(numbering, numberingId);

        return abstractNumbering is null ? null : ListNumbering.ClassifyAbstractNumbering(abstractNumbering);
    }

    /// <summary>
    /// Creates a relationship to an external target and returns its identifier.
    /// </summary>
    /// <param name="contextElement">Element the relationship will be referenced from.</param>
    /// <param name="target">External target.</param>
    internal string CreateExternalRelationship(OpenXmlElement contextElement, Uri target)
    {
        return ResolveOwningPart(contextElement).AddHyperlinkRelationship(target, isExternal: true).Id;
    }

    /// <summary>
    /// Adds an image part, fills it from <paramref name="content"/>, and returns its relationship id.
    /// </summary>
    /// <param name="contextElement">Element the image will be referenced from.</param>
    /// <param name="content">Image bytes.</param>
    /// <param name="partType">Package part type carrying the right content type.</param>
    internal string AddImagePart(OpenXmlElement contextElement, Stream content, PartTypeInfo partType)
    {
        var owningPart = ResolveOwningPart(contextElement);
        var imagePart = AddImagePartTo(owningPart, partType);
        imagePart.FeedData(content);

        return owningPart.GetIdOfPart(imagePart);
    }

    /// <summary>
    /// Adds an image part to whichever kind of part owns the content.
    /// </summary>
    /// <remarks>
    /// The SDK exposes <c>AddImagePart</c> per part type rather than on the container base, so the
    /// three parts that can hold body content are handled explicitly.
    /// </remarks>
    private static ImagePart AddImagePartTo(OpenXmlPartContainer owningPart, PartTypeInfo partType)
    {
        return owningPart switch
        {
            MainDocumentPart mainDocumentPart => mainDocumentPart.AddImagePart(partType),
            HeaderPart headerPart => headerPart.AddImagePart(partType),
            FooterPart footerPart => footerPart.AddImagePart(partType),
            _ => throw new NotSupportedException($"Cannot embed an image in a {owningPart.GetType().Name}."),
        };
    }

    /// <summary>
    /// Finds the package part that must own a relationship referenced from <paramref name="element"/>.
    /// </summary>
    /// <remarks>
    /// Relationships belong to the part that references them, not to the document as a whole. An image
    /// or hyperlink inside a header therefore has to be registered on that <c>HeaderPart</c>: registering
    /// it on the main document part produces an id the header cannot resolve, and Word reports the
    /// document as corrupt. The owning part is derived by walking up to the tree's root element rather
    /// than being passed down through every wrapper, so it stays correct wherever content ends up.
    /// </remarks>
    private OpenXmlPartContainer ResolveOwningPart(OpenXmlElement element)
    {
        var root = element;
        while (root.Parent is not null)
        {
            root = root.Parent;
        }

        return (root as OpenXmlPartRootElement)?.OpenXmlPart ?? MainDocumentPart;
    }

    /// <summary>
    /// Returns an identifier no other drawing in this document uses.
    /// </summary>
    /// <remarks>
    /// Drawing identifiers have to be unique within the document part, so the counter starts above
    /// whatever an opened document already contains rather than at one.
    /// </remarks>
    internal uint NextDrawingId()
    {
        if (!_drawingIdsSeeded)
        {
            _nextDrawingId = FindHighestDrawingId() + 1;
            _drawingIdsSeeded = true;
        }

        return _nextDrawingId++;
    }

    private uint FindHighestDrawingId()
    {
        // Headers and footers are scanned too, so that a document reopened after its logo was placed in
        // a header does not hand out an identifier that header already uses.
        var roots = new List<OpenXmlElement?> { MainDocumentPart.Document };
        roots.AddRange(MainDocumentPart.HeaderParts.Select(part => (OpenXmlElement?)part.Header));
        roots.AddRange(MainDocumentPart.FooterParts.Select(part => (OpenXmlElement?)part.Footer));

        var highestId = 0U;
        foreach (var root in roots)
        {
            if (root is null)
            {
                continue;
            }

            foreach (var properties in root.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>())
            {
                highestId = Math.Max(highestId, properties.Id?.Value ?? 0U);
            }
        }

        return highestId;
    }

    private WordLib.Styles GetOrCreateStyles()
    {
        var stylesPart = MainDocumentPart.StyleDefinitionsPart
            ?? MainDocumentPart.AddNewPart<StyleDefinitionsPart>();

        return stylesPart.Styles ??= new WordLib.Styles();
    }

    private WordLib.Numbering GetOrCreateNumbering()
    {
        var numberingPart = MainDocumentPart.NumberingDefinitionsPart
            ?? MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();

        return numberingPart.Numbering ??= new WordLib.Numbering();
    }

    private static bool ContainsStyle(WordLib.Styles styles, string styleId)
    {
        return styles.Elements<WordLib.Style>()
            .Any(style => string.Equals(style.StyleId?.Value, styleId, StringComparison.Ordinal));
    }

    private static int? FindExistingNumbering(WordLib.Numbering numbering, ListStyle style)
    {
        foreach (var instance in numbering.Elements<WordLib.NumberingInstance>())
        {
            if (instance.NumberID?.Value is not { } numberingId)
            {
                continue;
            }

            var abstractNumbering = FindAbstractNumbering(numbering, numberingId);
            if (abstractNumbering is not null && ListNumbering.ClassifyAbstractNumbering(abstractNumbering) == style)
            {
                return numberingId;
            }
        }

        return null;
    }

    private static WordLib.AbstractNum? FindAbstractNumbering(WordLib.Numbering numbering, int numberingId)
    {
        var instance = numbering.Elements<WordLib.NumberingInstance>()
            .FirstOrDefault(candidate => candidate.NumberID?.Value == numberingId);

        if (instance?.AbstractNumId?.Val?.Value is not { } abstractNumberingId)
        {
            return null;
        }

        return numbering.Elements<WordLib.AbstractNum>()
            .FirstOrDefault(candidate => candidate.AbstractNumberId?.Value == abstractNumberingId);
    }

    /// <summary>
    /// Inserts an abstract definition ahead of the concrete instances.
    /// </summary>
    /// <remarks>
    /// <c>CT_Numbering</c> is <c>numPicBullet*, abstractNum*, num*</c>, so appending an
    /// <c>w:abstractNum</c> after the first <c>w:num</c> produces an invalid numbering part.
    /// </remarks>
    private static void InsertAbstractNumbering(WordLib.Numbering numbering, WordLib.AbstractNum abstractNumbering)
    {
        var firstInstance = numbering.GetFirstChild<WordLib.NumberingInstance>();
        if (firstInstance is null)
        {
            numbering.AppendChild(abstractNumbering);
            return;
        }

        numbering.InsertBefore(abstractNumbering, firstInstance);
    }

    private static int NextAbstractNumberingId(WordLib.Numbering numbering)
    {
        var highestId = -1;
        foreach (var abstractNumbering in numbering.Elements<WordLib.AbstractNum>())
        {
            highestId = Math.Max(highestId, abstractNumbering.AbstractNumberId?.Value ?? -1);
        }

        return highestId + 1;
    }

    private static int NextNumberingId(WordLib.Numbering numbering)
    {
        // Numbering ids start at 1, because 0 is reserved for "no list".
        var highestId = 0;
        foreach (var instance in numbering.Elements<WordLib.NumberingInstance>())
        {
            highestId = Math.Max(highestId, instance.NumberID?.Value ?? 0);
        }

        return highestId + 1;
    }
}
