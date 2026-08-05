namespace OfficeDocuments.Word.Enums;

/// <summary>
/// Vertical position of text relative to the baseline.
/// </summary>
public enum TextVerticalPosition
{
    /// <summary>
    /// On the baseline. Clears a superscript or subscript inherited from a style.
    /// </summary>
    Baseline,
    /// <summary>
    /// Raised and reduced, as in <c>m²</c>.
    /// </summary>
    Superscript,
    /// <summary>
    /// Lowered and reduced, as in <c>H₂O</c>.
    /// </summary>
    Subscript,
}
