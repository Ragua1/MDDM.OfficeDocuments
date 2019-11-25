using DocumentFormat.OpenXml;

namespace OfficeDocumentsApi.Excel.Enums
{
    /// <summary>
    /// Values of HorizontalAlignment
    /// </summary>
    public enum HorizontalAlignmentValues
    {
        /// <summary>
        /// General value
        /// </summary>
        [EnumString("general")] General,
        /// <summary>
        /// left value
        /// </summary>
        [EnumString("left")] Left,
        /// <summary>
        /// Center value
        /// </summary>
        [EnumString("center")] Center,
        /// <summary>
        /// Right value
        /// </summary>
        [EnumString("right")] Right,
        /// <summary>
        /// Fill value
        /// </summary>
        [EnumString("fill")] Fill,
        /// <summary>
        /// Justify value
        /// </summary>
        [EnumString("justify")] Justify,
        /// <summary>
        /// CenterContinuous value
        /// </summary>
        [EnumString("centerContinuous")] CenterContinuous,
        /// <summary>
        /// Distributed value
        /// </summary>
        [EnumString("distributed")] Distributed,
    }
}