using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;
using OfficeDocuments.Word.Interfaces;

namespace OfficeDocuments.Word;

/// <summary>
/// Entry point of the Word module: owns the <see cref="WordprocessingDocument"/> and its lifetime.
/// </summary>
public class Wordprocessing : IWordprocessing
{
    /// <summary>
    /// Opens with the SDK's auto-save turned off, so that this class decides when the package is written.
    /// </summary>
    /// <remarks>
    /// The SDK saves modifications when the package is disposed unless auto-save is off, and it makes
    /// that decision at open time rather than at close time. Left at its default, the
    /// <c>saveDocument: false</c> argument of <see cref="Close"/> could not do anything: it skipped an
    /// explicit save that disposal then performed anyway, so discarding changes quietly kept them. With
    /// auto-save off, the only thing that writes the package is <see cref="Close"/> deciding to.
    /// </remarks>
    private static OpenSettings ManualSaveSettings => new() { AutoSave = false };

    private readonly WordprocessingDocument _document;
    private readonly bool _isEditable;
    private readonly Dictionary<(bool IsHeader, HeaderFooterKind Kind), DataClasses.HeaderFooter> _headersAndFooters = [];
    private DataClasses.Body? _body;
    private DataClasses.DocumentContext? _context;
    private bool _closed;

    /// <summary>
    /// Creates a document in <paramref name="stream"/>, or opens the one it already holds.
    /// </summary>
    /// <param name="stream">Stream holding the document package.</param>
    /// <param name="createNew"><see langword="true"/> to create a new document.</param>
    /// <param name="isEditable">
    /// <see langword="false"/> to open an existing document for reading only, which also stops
    /// <see cref="Close"/> and <see cref="Dispose"/> from writing it back.
    /// </param>
    public Wordprocessing(Stream stream, bool createNew, bool isEditable = true) :
        this(createNew
                ? WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false)
                : WordprocessingDocument.Open(stream, isEditable, ManualSaveSettings),
            createNew,
            createNew || isEditable)
    { }

    /// <summary>
    /// Creates the document at <paramref name="filePath"/>, or opens the existing one.
    /// </summary>
    /// <param name="filePath">Path of the document.</param>
    /// <param name="createNew"><see langword="true"/> to create a new document.</param>
    /// <param name="isEditable">
    /// <see langword="false"/> to open an existing document for reading only, which also stops
    /// <see cref="Close"/> and <see cref="Dispose"/> from writing it back.
    /// </param>
    public Wordprocessing(string filePath, bool createNew, bool isEditable = true) :
        this(createNew
                ? WordprocessingDocument.Create(Path.GetFullPath(filePath), WordprocessingDocumentType.Document, autoSave: false)
                : WordprocessingDocument.Open(Path.GetFullPath(filePath), isEditable, ManualSaveSettings),
            createNew,
            createNew || isEditable)
    { }

    /// <summary>
    /// Takes ownership of an already-opened package. The public constructors funnel into this one.
    /// </summary>
    /// <param name="document">The package to wrap. This instance disposes it.</param>
    /// <param name="createNew">
    /// <see langword="true"/> when the package was just created and still needs its main document
    /// part and body.
    /// </param>
    /// <param name="isEditable">
    /// <see langword="false"/> to leave the file untouched on close, so reading a document does not
    /// rewrite its bytes.
    /// </param>
    /// <exception cref="ArgumentNullException"><paramref name="document"/> is <see langword="null"/>.</exception>
    protected internal Wordprocessing(WordprocessingDocument document, bool createNew, bool isEditable = true)
    {
        ArgumentNullException.ThrowIfNull(document);

        _document = document;
        _isEditable = isEditable;

        if (createNew)
        {
            document.AddMainDocumentPart().Document = new Document();
        }
        else if (document.MainDocumentPart?.Document is null)
        {
            throw new InvalidOperationException("The document does not contain a main document part with a root document element.");
        }
    }

    /// <inheritdoc />
    public IBody GetBody()
    {
        ObjectDisposedException.ThrowIf(_closed, this);

        return _body ??= CreateBody();
    }

    private DataClasses.Body CreateBody()
    {
        var mainDocumentPart = _document.MainDocumentPart
            ?? throw new InvalidOperationException("The word document does not contain a main document part.");
        var documentElement = mainDocumentPart.Document
            ?? throw new InvalidOperationException("The word document does not contain a root document element.");

        // A document created through the SDK has no w:body until something needs one.
        var bodyElement = documentElement.Body ?? documentElement.AppendChild(new Body());

        // One context per document, shared with headers and footers, so styles, list numbering, and
        // drawing identifiers are tracked once rather than per container.
        _context ??= new DataClasses.DocumentContext(mainDocumentPart);

        return new DataClasses.Body(bodyElement, _context);
    }

    /// <inheritdoc />
    public IHeaderFooter AddHeader(HeaderFooterKind kind = HeaderFooterKind.Default)
        => AddHeaderFooter(kind, isHeader: true);

    /// <inheritdoc />
    public IHeaderFooter AddFooter(HeaderFooterKind kind = HeaderFooterKind.Default)
        => AddHeaderFooter(kind, isHeader: false);

    /// <inheritdoc />
    public IReadOnlyList<IHeaderFooter> HeadersAndFooters
    {
        get
        {
            ObjectDisposedException.ThrowIf(_closed, this);

            return DiscoverHeadersAndFooters();
        }
    }

    /// <inheritdoc />
    public int ReplaceText(string oldValue, string newValue, StringComparison comparison = StringComparison.Ordinal)
    {
        ArgumentException.ThrowIfNullOrEmpty(oldValue);
        ArgumentNullException.ThrowIfNull(newValue);
        ObjectDisposedException.ThrowIf(_closed, this);

        var replaced = GetBodyInternal().ReplaceText(oldValue, newValue, comparison);

        foreach (var container in DiscoverHeadersAndFooters())
        {
            replaced += container.ReplaceText(oldValue, newValue, comparison);
        }

        return replaced;
    }

    /// <summary>
    /// Reads the headers and footers the section references, wrapping each one once.
    /// </summary>
    /// <remarks>
    /// Derived from the document rather than from what this instance created, because a document opened
    /// from disk arrives with headers and footers already in it and reporting an empty list for those
    /// would make every read and template scenario miss them. Wrappers are cached per kind, so a caller
    /// holding on to the result of <see cref="AddHeader"/> keeps the same instance.
    /// </remarks>
    private List<DataClasses.HeaderFooter> DiscoverHeadersAndFooters()
    {
        var mainDocumentPart = RequireMainDocumentPart();
        var sectionProperties = GetBodyInternal().GetOrCreateSectionProperties();
        var found = new List<DataClasses.HeaderFooter>();

        foreach (var child in sectionProperties.ChildElements)
        {
            var (isHeader, referenceType, relationshipId) = child switch
            {
                HeaderReference header => (true, header.Type?.Value ?? HeaderFooterValues.Default, header.Id?.Value),
                FooterReference footer => (false, footer.Type?.Value ?? HeaderFooterValues.Default, footer.Id?.Value),
                _ => (false, HeaderFooterValues.Default, null),
            };

            if (relationshipId is null || ToKind(referenceType) is not { } kind)
            {
                continue;
            }

            var container = isHeader
                ? (mainDocumentPart.GetPartById(relationshipId) as HeaderPart)?.Header
                : (OpenXmlCompositeElement?)(mainDocumentPart.GetPartById(relationshipId) as FooterPart)?.Footer;

            if (container is not null)
            {
                found.Add(GetOrCreateHeaderFooter(container, kind, isHeader));
            }
        }

        return found;
    }

    private DataClasses.HeaderFooter GetOrCreateHeaderFooter(
        OpenXmlCompositeElement container,
        HeaderFooterKind kind,
        bool isHeader)
    {
        var key = (isHeader, kind);
        if (_headersAndFooters.TryGetValue(key, out var existing))
        {
            return existing;
        }

        var created = new DataClasses.HeaderFooter(container, GetContext(), kind, isHeader);
        _headersAndFooters[key] = created;

        return created;
    }

    /// <inheritdoc />
    public PageSetup PageSetup
    {
        get
        {
            ObjectDisposedException.ThrowIf(_closed, this);

            return PageSetupMapper.Read(GetBodyInternal().GetOrCreateSectionProperties());
        }
    }

    /// <inheritdoc />
    public IWordprocessing ApplyPageSetup(PageSetup setup)
    {
        ArgumentNullException.ThrowIfNull(setup);
        ObjectDisposedException.ThrowIf(_closed, this);

        PageSetupMapper.Apply(GetBodyInternal().GetOrCreateSectionProperties(), setup);

        return this;
    }

    /// <inheritdoc />
    public DocumentMetadata Metadata
    {
        get
        {
            ObjectDisposedException.ThrowIf(_closed, this);

            var properties = _document.PackageProperties;

            return new DocumentMetadata
            {
                Title = properties.Title,
                Subject = properties.Subject,
                Author = properties.Creator,
                Keywords = properties.Keywords,
                Description = properties.Description,
                Category = properties.Category,
                LastModifiedBy = properties.LastModifiedBy,
                Created = properties.Created,
                Modified = properties.Modified,
            };
        }
    }

    /// <inheritdoc />
    public IWordprocessing SetMetadata(DocumentMetadata metadata)
    {
        ArgumentNullException.ThrowIfNull(metadata);
        ObjectDisposedException.ThrowIf(_closed, this);

        var properties = _document.PackageProperties;

        if (metadata.Title is not null) properties.Title = metadata.Title;
        if (metadata.Subject is not null) properties.Subject = metadata.Subject;
        if (metadata.Author is not null) properties.Creator = metadata.Author;
        if (metadata.Keywords is not null) properties.Keywords = metadata.Keywords;
        if (metadata.Description is not null) properties.Description = metadata.Description;
        if (metadata.Category is not null) properties.Category = metadata.Category;
        if (metadata.LastModifiedBy is not null) properties.LastModifiedBy = metadata.LastModifiedBy;
        if (metadata.Created is { } created) properties.Created = created.UtcDateTime;
        if (metadata.Modified is { } modified) properties.Modified = modified.UtcDateTime;

        return this;
    }

    private DataClasses.HeaderFooter AddHeaderFooter(HeaderFooterKind kind, bool isHeader)
    {
        ObjectDisposedException.ThrowIf(_closed, this);

        var mainDocumentPart = RequireMainDocumentPart();
        var sectionProperties = GetBodyInternal().GetOrCreateSectionProperties();
        var referenceType = ToOpenXml(kind);

        // A document opened from disk may already reference a header of this kind; reusing its part
        // keeps the existing content instead of orphaning it behind a second reference.
        var container = FindExistingHeaderFooter(mainDocumentPart, sectionProperties, referenceType, isHeader)
                        ?? CreateHeaderFooter(mainDocumentPart, sectionProperties, referenceType, isHeader);

        EnableHeaderFooterKind(mainDocumentPart, sectionProperties, kind);

        return GetOrCreateHeaderFooter(container, kind, isHeader);
    }

    private static OpenXmlCompositeElement? FindExistingHeaderFooter(
        MainDocumentPart mainDocumentPart,
        SectionProperties sectionProperties,
        HeaderFooterValues referenceType,
        bool isHeader)
    {
        var relationshipId = isHeader
            ? sectionProperties.Elements<HeaderReference>()
                .FirstOrDefault(reference => (reference.Type?.Value ?? HeaderFooterValues.Default) == referenceType)?.Id?.Value
            : sectionProperties.Elements<FooterReference>()
                .FirstOrDefault(reference => (reference.Type?.Value ?? HeaderFooterValues.Default) == referenceType)?.Id?.Value;

        if (relationshipId is null)
        {
            return null;
        }

        return isHeader
            ? (mainDocumentPart.GetPartById(relationshipId) as HeaderPart)?.Header
            : (mainDocumentPart.GetPartById(relationshipId) as FooterPart)?.Footer;
    }

    private static OpenXmlCompositeElement CreateHeaderFooter(
        MainDocumentPart mainDocumentPart,
        SectionProperties sectionProperties,
        HeaderFooterValues referenceType,
        bool isHeader)
    {
        if (isHeader)
        {
            var part = mainDocumentPart.AddNewPart<HeaderPart>();
            var header = part.Header = new Header();

            DataClasses.SectionPropertiesOrderer.Insert(
                sectionProperties,
                new HeaderReference { Id = mainDocumentPart.GetIdOfPart(part), Type = referenceType });

            return header;
        }

        var footerPart = mainDocumentPart.AddNewPart<FooterPart>();
        var footer = footerPart.Footer = new Footer();

        DataClasses.SectionPropertiesOrderer.Insert(
            sectionProperties,
            new FooterReference { Id = mainDocumentPart.GetIdOfPart(footerPart), Type = referenceType });

        return footer;
    }

    /// <summary>
    /// Turns on the switch a first-page or even-page header needs to be displayed.
    /// </summary>
    /// <remarks>
    /// Both are easy to get wrong because the document is perfectly valid without them — the header
    /// simply never appears. A first-page header requires <c>w:titlePg</c> on the section, and an
    /// even-page header requires <c>w:evenAndOddHeaders</c> in the document settings.
    /// </remarks>
    private static void EnableHeaderFooterKind(
        MainDocumentPart mainDocumentPart,
        SectionProperties sectionProperties,
        HeaderFooterKind kind)
    {
        switch (kind)
        {
            case HeaderFooterKind.First:
                DataClasses.SectionPropertiesOrderer.GetOrCreate(sectionProperties, () => new TitlePage());
                break;

            case HeaderFooterKind.Even:
                var settingsPart = mainDocumentPart.DocumentSettingsPart
                                   ?? mainDocumentPart.AddNewPart<DocumentSettingsPart>();
                var settings = settingsPart.Settings ??= new Settings();

                if (settings.GetFirstChild<EvenAndOddHeaders>() is null)
                {
                    settings.AppendChild(new EvenAndOddHeaders());
                }

                break;
        }
    }

    private static HeaderFooterValues ToOpenXml(HeaderFooterKind kind)
    {
        return kind switch
        {
            HeaderFooterKind.Default => HeaderFooterValues.Default,
            HeaderFooterKind.First => HeaderFooterValues.First,
            HeaderFooterKind.Even => HeaderFooterValues.Even,
            _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unsupported header or footer kind."),
        };
    }

    /// <summary>
    /// Maps a reference type found in a document back to the kind this library exposes, or
    /// <see langword="null"/> for one it does not model.
    /// </summary>
    /// <remarks>
    /// Returns <see langword="null"/> rather than throwing, because this reads whatever a document
    /// happens to contain. Refusing to list a document's headers over one reference the library does not
    /// understand would be worse than leaving that one out.
    /// </remarks>
    private static HeaderFooterKind? ToKind(HeaderFooterValues referenceType)
    {
        if (referenceType == HeaderFooterValues.Default)
        {
            return HeaderFooterKind.Default;
        }

        if (referenceType == HeaderFooterValues.First)
        {
            return HeaderFooterKind.First;
        }

        return referenceType == HeaderFooterValues.Even ? HeaderFooterKind.Even : null;
    }

    private MainDocumentPart RequireMainDocumentPart()
    {
        return _document.MainDocumentPart
            ?? throw new InvalidOperationException("The word document does not contain a main document part.");
    }

    private DataClasses.Body GetBodyInternal()
    {
        return _body ??= CreateBody();
    }

    private DataClasses.DocumentContext GetContext()
    {
        // Creating the body is what creates the context, and a header is meaningless without a body.
        GetBodyInternal();

        return _context ?? throw new InvalidOperationException("The document context was not initialized.");
    }

    #region IDisposable implementation

    /// <summary>
    /// Saves and closes the document.
    /// </summary>
    /// <remarks>
    /// Idempotent by design, so the documented <c>using</c> plus explicit <c>Close()</c> pattern
    /// works: the second call returns instead of saving an already-disposed package.
    /// </remarks>
    /// <param name="saveDocument">
    /// <see langword="false"/> to discard changes. Ignored for a document opened read-only.
    /// </param>
    public void Close(bool saveDocument = true)
    {
        if (_closed)
        {
            return;
        }

        // Set before disposing so a failure cannot leave the instance looking reusable.
        _closed = true;

        if (_isEditable && saveDocument)
        {
            _document.Save();
        }

        _document.Dispose();
    }

    /// <summary>
    /// Saves and releases the document, unless it was already closed.
    /// </summary>
    public void Dispose()
    {
        Close();
        GC.SuppressFinalize(this);
    }

    #endregion
}
