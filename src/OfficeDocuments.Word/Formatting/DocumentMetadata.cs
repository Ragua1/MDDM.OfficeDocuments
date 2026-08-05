namespace OfficeDocuments.Word.Formatting;

/// <summary>
/// The document's core properties: what Word shows under File, Info.
/// </summary>
/// <remarks>
/// <para>
/// As with the other records, <see langword="null"/> means "leave this alone". To clear a property that
/// is already set, assign an empty string.
/// </para>
/// <para>
/// These are the standard package properties, so they are also what a document management system, a
/// search index, or Windows Explorer reads.
/// </para>
/// </remarks>
public sealed record DocumentMetadata
{
    /// <summary>
    /// Document title. Distinct from the file name and from a <c>Title</c>-styled paragraph.
    /// </summary>
    public string? Title { get; init; }

    /// <summary>
    /// What the document is about.
    /// </summary>
    public string? Subject { get; init; }

    /// <summary>
    /// Who wrote it. Stored as the core property <c>dc:creator</c>, which Word labels "Author".
    /// </summary>
    public string? Author { get; init; }

    /// <summary>
    /// Search keywords, conventionally separated by semicolons or commas.
    /// </summary>
    public string? Keywords { get; init; }

    /// <summary>
    /// Longer description. Stored as <c>dc:description</c>, which Word labels "Comments".
    /// </summary>
    public string? Description { get; init; }

    /// <summary>
    /// Classification of the document, for example a document type.
    /// </summary>
    public string? Category { get; init; }

    /// <summary>
    /// Who saved it last.
    /// </summary>
    public string? LastModifiedBy { get; init; }

    /// <summary>
    /// When the document was created.
    /// </summary>
    public DateTimeOffset? Created { get; init; }

    /// <summary>
    /// When the document was last changed.
    /// </summary>
    public DateTimeOffset? Modified { get; init; }

    /// <summary>
    /// <see langword="true"/> when no property is set, so applying it would write nothing.
    /// </summary>
    public bool IsEmpty =>
        Title is null
        && Subject is null
        && Author is null
        && Keywords is null
        && Description is null
        && Category is null
        && LastModifiedBy is null
        && Created is null
        && Modified is null;

    /// <summary>
    /// Layers <paramref name="overrides"/> on top of this metadata: every property the argument sets
    /// wins, and the ones it leaves unset keep this value.
    /// </summary>
    /// <param name="overrides">Metadata whose set properties take precedence. May be <see langword="null"/>.</param>
    /// <returns>The combined metadata. Neither input is modified.</returns>
    public DocumentMetadata Merge(DocumentMetadata? overrides)
    {
        if (overrides is null || overrides.IsEmpty)
        {
            return this;
        }

        return new DocumentMetadata
        {
            Title = overrides.Title ?? Title,
            Subject = overrides.Subject ?? Subject,
            Author = overrides.Author ?? Author,
            Keywords = overrides.Keywords ?? Keywords,
            Description = overrides.Description ?? Description,
            Category = overrides.Category ?? Category,
            LastModifiedBy = overrides.LastModifiedBy ?? LastModifiedBy,
            Created = overrides.Created ?? Created,
            Modified = overrides.Modified ?? Modified,
        };
    }
}
