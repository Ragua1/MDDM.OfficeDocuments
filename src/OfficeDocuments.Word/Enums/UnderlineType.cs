namespace OfficeDocuments.Word.Enums;

/// <summary>
/// Underline styles available on a text run.
/// </summary>
public enum UnderlineType
{
    /// <summary>
    /// No underline. Written explicitly, so it can switch off an underline inherited from a style.
    /// </summary>
    None,
    /// <summary>
    /// A single continuous line.
    /// </summary>
    Single,
    /// <summary>
    /// Two continuous lines.
    /// </summary>
    Double,
    /// <summary>
    /// A single thick line.
    /// </summary>
    Thick,
    /// <summary>
    /// A dotted line.
    /// </summary>
    Dotted,
    /// <summary>
    /// A dashed line.
    /// </summary>
    Dash,
    /// <summary>
    /// A wavy line.
    /// </summary>
    Wave,
    /// <summary>
    /// A single line under each word, skipping the spaces between them.
    /// </summary>
    Words,
}
