namespace OpenXmlApi.Interfaces
{
    /// <summary>
    /// Interface of Base class
    /// </summary>
    public interface IBase
    {
        /// <summary>
        /// Instance of worksheet
        /// </summary>
        IWorksheet Worksheet { get; }
        /// <summary>
        /// Instance of custom style
        /// </summary>
        IStyle Style { get; }

        /// <summary>
        /// Add custom styles 
        /// </summary>
        /// <param name="styles">Custom styles</param>
        /// <returns>Style of object</returns>
        IStyle AddStyle(params IStyle[] styles);
    }
}