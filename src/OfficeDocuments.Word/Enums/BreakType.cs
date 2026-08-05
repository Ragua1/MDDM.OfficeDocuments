namespace OfficeDocuments.Word.Enums;

/// <summary>
/// The kind of break <see cref="Interfaces.IParagraph.AddBreak"/> inserts.
/// </summary>
public enum BreakType
{
    /// <summary>
    /// Starts a new page.
    /// </summary>
    Page,

    /// <summary>
    /// Starts a new column, in a section laid out in columns.
    /// </summary>
    Column,

    /// <summary>
    /// A line break inside the same paragraph. This is what a newline in the text passed to
    /// <see cref="Interfaces.IParagraph.AddText(string)"/> becomes, and it reads back as
    /// <c>\n</c>.
    /// </summary>
    TextWrapping,
}
