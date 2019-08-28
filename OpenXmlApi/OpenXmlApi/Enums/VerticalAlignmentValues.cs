using DocumentFormat.OpenXml;

namespace OpenXmlApi.Emums
{
    /// <summary>
    /// Values of VerticalAlignment
    /// </summary>
    public enum VerticalAlignmentValues
    {
        /// <summary>
        /// Top value
        /// </summary>
        [EnumString("top")] Top,
        /// <summary>
        /// Center value
        /// </summary>
        [EnumString("center")] Center,
        /// <summary>
        /// Bottom value
        /// </summary>
        [EnumString("bottom")] Bottom,
        /// <summary>
        /// Justify value
        /// </summary>
        [EnumString("justify")] Justify,
        /// <summary>
        /// Distributed value
        /// </summary>
        [EnumString("distributed")] Distributed,
    }
}