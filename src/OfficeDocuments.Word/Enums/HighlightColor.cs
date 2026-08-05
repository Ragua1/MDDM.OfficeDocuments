namespace OfficeDocuments.Word.Enums;

/// <summary>
/// Text highlight colours. WordprocessingML allows only this fixed palette for
/// <c>w:highlight</c>, not arbitrary colour values.
/// </summary>
public enum HighlightColor
{
    /// <summary>No highlight. Clears a highlight inherited from a style.</summary>
    None,
    /// <summary>Yellow.</summary>
    Yellow,
    /// <summary>Bright green.</summary>
    Green,
    /// <summary>Cyan.</summary>
    Cyan,
    /// <summary>Magenta.</summary>
    Magenta,
    /// <summary>Blue.</summary>
    Blue,
    /// <summary>Red.</summary>
    Red,
    /// <summary>Dark blue.</summary>
    DarkBlue,
    /// <summary>Dark cyan.</summary>
    DarkCyan,
    /// <summary>Dark green.</summary>
    DarkGreen,
    /// <summary>Dark magenta.</summary>
    DarkMagenta,
    /// <summary>Dark red.</summary>
    DarkRed,
    /// <summary>Dark yellow.</summary>
    DarkYellow,
    /// <summary>Dark grey.</summary>
    DarkGray,
    /// <summary>Light grey.</summary>
    LightGray,
    /// <summary>Black.</summary>
    Black,
    /// <summary>White.</summary>
    White,
}
