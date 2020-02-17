using DocumentFormat.OpenXml;

namespace OfficeDocumentsApi.Emums
{
    /// <summary>
    /// Values of Underline
    /// </summary>
    public enum UnderlineValues
    {
        /// <summary>
        /// Single value
        /// </summary>
        [EnumString("single")] Single,
        /// <summary>
        /// Double value
        /// </summary>
        [EnumString("double")] Double,
        /// <summary>
        /// SingleAccounting value
        /// </summary>
        [EnumString("singleAccounting")] SingleAccounting,
        /// <summary>
        /// DoubleAccounting value
        /// </summary>
        [EnumString("doubleAccounting")] DoubleAccounting,
        /// <summary>
        /// None value
        /// </summary>
        [EnumString("none")] None,
    }
}