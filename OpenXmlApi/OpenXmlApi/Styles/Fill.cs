using System.Drawing;

namespace OpenXmlApi.Styles
{
    /// <summary>
    /// Class of Fill
    /// </summary>
    public class Fill
    {
        /// <summary>
        /// Instance of Fill element
        /// </summary>
        public DocumentFormat.OpenXml.Spreadsheet.Fill Element { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Fill"/> class.
        /// </summary>
        /// <param name="fill">Spreadsheet fill.</param>
        public Fill(DocumentFormat.OpenXml.Spreadsheet.Fill fill = null)
        {
            this.Element = fill ?? new DocumentFormat.OpenXml.Spreadsheet.Fill
            {
                PatternFill = new DocumentFormat.OpenXml.Spreadsheet.PatternFill()
            };
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="Fill"/> class.
        /// </summary>
        /// <param name="foregroundColor">Color of the foreground.</param>
        /// <param name="pattern">The pattern.</param>
        public Fill(Color foregroundColor, DocumentFormat.OpenXml.Spreadsheet.PatternValues pattern = DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
            : this(Utils.ArgbHexConverter(foregroundColor), pattern)
        { }
        /// <summary>
        /// Initializes a new instance of the <see cref="Fill"/> class.
        /// </summary>
        /// <param name="foregroundColor">Color of the foreground.</param>
        /// <param name="pattern">The pattern.</param>
        public Fill(string foregroundColor, DocumentFormat.OpenXml.Spreadsheet.PatternValues pattern = DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
        {
            this.Element = new DocumentFormat.OpenXml.Spreadsheet.Fill
            {
                PatternFill = new DocumentFormat.OpenXml.Spreadsheet.PatternFill
                {
                    ForegroundColor = new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor() { Rgb = foregroundColor },
                    PatternType = pattern,
                }
            };
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="Fill"/> class.
        /// </summary>
        /// <param name="backgroundColor">Color of the background.</param>
        /// <param name="foregroundColor">Color of the foreground.</param>
        /// <param name="pattern">The pattern.</param>
        public Fill(Color backgroundColor, Color foregroundColor, DocumentFormat.OpenXml.Spreadsheet.PatternValues pattern = DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
        {
            this.Element = new DocumentFormat.OpenXml.Spreadsheet.Fill
            {
                PatternFill = new DocumentFormat.OpenXml.Spreadsheet.PatternFill
                {
                    BackgroundColor = new DocumentFormat.OpenXml.Spreadsheet.BackgroundColor { Rgb = Utils.ArgbHexConverter(backgroundColor) },
                    ForegroundColor = new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor { Rgb = Utils.ArgbHexConverter(foregroundColor) },
                    PatternType = pattern,
                }
            };
        }

        /// <summary>
        /// Compare content with 'fill'
        /// </summary>
        /// <param name="fill">Spreadsheet fill for compare</param>
        public bool IsContentSame(DocumentFormat.OpenXml.Spreadsheet.Fill fill)
        {
            return Utils.CompareXml(fill.OuterXml, this.Element.OuterXml);
        }
    }
}