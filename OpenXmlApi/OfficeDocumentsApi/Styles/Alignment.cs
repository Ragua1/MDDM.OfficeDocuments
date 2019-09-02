using OfficeDocumentsApi.Emums;

namespace OfficeDocumentsApi.Styles
{
    /// <summary>
    /// Class of Alignment
    /// </summary>
    public class Alignment
    {
        /// <summary>
        /// Instance of Alignment element
        /// </summary>
        public DocumentFormat.OpenXml.Spreadsheet.Alignment Element { get; }
        /// <summary>
        /// Sets the horizontal.
        /// </summary>
        public HorizontalAlignmentValues Horizontal
        {
            set { Element.Horizontal = (DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues)value; }
        }
        /// <summary>
        /// Sets the vertical.
        /// </summary>
        public VerticalAlignmentValues Vertical
        {
            set { Element.Vertical = (DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues)value; }
        }
        /// <summary>
        /// Sets a value indicating whether [justify last line].
        /// </summary>
        public bool JustifyLastLine
        {
            set { Element.JustifyLastLine = value; }
        }
        /// <summary>
        /// Sets a value indicating whether [wrap text].
        /// </summary>
        public bool WrapText
        {
            set { Element.WrapText = value; }
        }
        /// <summary>
        /// Sets a value indicating whether [shrink to fit].
        /// </summary>
        public bool ShrinkToFit
        {
            set { Element.ShrinkToFit = value; }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Alignment"/> class.
        /// </summary>
        /// <param name="alignment">Spreadsheet alignment</param>
        public Alignment(DocumentFormat.OpenXml.Spreadsheet.Alignment alignment = null)
        {
            Element = alignment ?? new DocumentFormat.OpenXml.Spreadsheet.Alignment();
        }

        /// <summary>
        /// Compare content with 'alignment'
        /// </summary>
        /// <param name="alignment">Spreadsheet alignment for compare</param>
        public bool IsContentSame(DocumentFormat.OpenXml.Spreadsheet.Alignment alignment)
        {
            return Utils.CompareXml(alignment.OuterXml, Element.OuterXml);
        }
    }
}