namespace OfficeDocuments.Word.Enums;

/// <summary>
/// Kind of list a paragraph belongs to.
/// </summary>
public enum ListStyle
{
    /// <summary>
    /// Not a list item. Removes a paragraph from a list it was in.
    /// </summary>
    None,
    /// <summary>
    /// Unordered list marked with bullet characters.
    /// </summary>
    Bullet,
    /// <summary>
    /// Ordered list numbered <c>1.</c>, <c>2.</c>, <c>3.</c> and so on.
    /// </summary>
    Number,
}
