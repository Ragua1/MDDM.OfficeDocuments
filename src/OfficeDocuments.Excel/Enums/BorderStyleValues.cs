using System.Runtime.Serialization;
using DocumentFormat.OpenXml;

namespace OfficeDocuments.Excel.Enums
{
    /// <summary>
    /// Values for BorderStyle
    /// </summary>
    public enum BorderStyleValues
    {
        /// <summary>
        /// None value
        /// </summary>
        [EnumMember(Value = "none")]
        None,
        /// <summary>
        /// Thin value
        /// </summary>
        [EnumMember(Value = "thin")]
        Thin,
        /// <summary>
        /// Medium value
        /// </summary>
        [EnumMember(Value = "medium")]
        Medium,
        /// <summary>
        /// Dashed value
        /// </summary>
        [EnumMember(Value = "dashed")]
        Dashed,
        /// <summary>
        /// Dotted value
        /// </summary>
        [EnumMember(Value = "dotted")]
        Dotted,
        /// <summary>
        /// Thick value
        /// </summary>
        [EnumMember(Value = "thick")]
        Thick,
        /// <summary>
        /// Double value
        /// </summary>
        [EnumMember(Value = "double")]
        Double,
        /// <summary>
        /// Hair value
        /// </summary>
        [EnumMember(Value = "hair")]
        Hair,
        /// <summary>
        /// MediumDashed value
        /// </summary>
        [EnumMember(Value = "mediumDashed")]
        MediumDashed,
        /// <summary>
        /// DashDot value
        /// </summary>
        [EnumMember(Value = "dashDot")]
        DashDot,
        /// <summary>
        /// MediumDashDot value
        /// </summary>
        [EnumMember(Value = "mediumDashDot")]
        MediumDashDot,
        /// <summary>
        /// DashDotDot value
        /// </summary>
        [EnumMember(Value = "dashDotDot")]
        DashDotDot,
        /// <summary>
        /// MediumDashDotDot value
        /// </summary>
        [EnumMember(Value = "mediumDashDotDot")]
        MediumDashDotDot,
        /// <summary>
        /// SlantDashDot value
        /// </summary>
        [EnumMember(Value = "slantDashDot")]
        SlantDashDot,
    }
}