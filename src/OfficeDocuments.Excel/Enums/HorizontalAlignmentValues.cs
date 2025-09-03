using System.Runtime.Serialization;

namespace OfficeDocuments.Excel.Enums;

/// <summary>
/// Values of HorizontalAlignment
/// </summary>
public enum HorizontalAlignmentValues
{
    /// <summary>
    /// General value
    /// </summary>
    [EnumMember(Value = "general")]
    General,
    /// <summary>
    /// left value
    /// </summary>
    [EnumMember(Value = "left")]
    Left,
    /// <summary>
    /// Center value
    /// </summary>
    [EnumMember(Value = "center")]
    Center,
    /// <summary>
    /// Right value
    /// </summary>
    [EnumMember(Value = "right")]
    Right,
    /// <summary>
    /// Fill value
    /// </summary>
    [EnumMember(Value = "fill")]
    Fill,
    /// <summary>
    /// Justify value
    /// </summary>
    [EnumMember(Value = "justify")]
    Justify,
    /// <summary>
    /// CenterContinuous value
    /// </summary>
    [EnumMember(Value = "centerContinuous")]
    CenterContinuous,
    /// <summary>
    /// Distributed value
    /// </summary>
    [EnumMember(Value = "distributed")]
    Distributed,
}