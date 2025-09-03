using System.Runtime.Serialization;

namespace OfficeDocuments.Excel.Enums;

/// <summary>
/// Values of Underline
/// </summary>
public enum UnderlineValues
{
    /// <summary>
    /// Single value
    /// </summary>
    [EnumMember(Value = "single")]
    Single,
    /// <summary>
    /// Double value
    /// </summary>
    [EnumMember(Value = "double")]
    Double,
    /// <summary>
    /// SingleAccounting value
    /// </summary>
    [EnumMember(Value = "singleAccounting")]
    SingleAccounting,
    /// <summary>
    /// DoubleAccounting value
    /// </summary>
    [EnumMember(Value = "doubleAccounting")]
    DoubleAccounting,
    /// <summary>
    /// None value
    /// </summary>
    [EnumMember(Value = "none")]
    None,
}