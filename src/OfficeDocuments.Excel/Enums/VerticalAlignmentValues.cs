using System.Runtime.Serialization;
using DocumentFormat.OpenXml;

namespace OfficeDocuments.Excel.Enums
{
    /// <summary>
    /// Values of VerticalAlignment
    /// </summary>
    public enum VerticalAlignmentValues
    {
        /// <summary>
        /// Top value
        /// </summary>
        [EnumMember(Value = "top")]
        Top,
        /// <summary>
        /// Center value
        /// </summary>
        [EnumMember(Value = "center")]
        Center,
        /// <summary>
        /// Bottom value
        /// </summary>
        [EnumMember(Value = "bottom")]
        Bottom,
        /// <summary>
        /// Justify value
        /// </summary>
        [EnumMember(Value = "justify")]
        Justify,
        /// <summary>
        /// Distributed value
        /// </summary>
        [EnumMember(Value = "distributed")]
        Distributed,
    }
}