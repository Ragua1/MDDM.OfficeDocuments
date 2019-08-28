
using OpenXmlApi.Emums;

namespace OpenXmlApi.Styles
{
    /// <summary>
    /// Class of Border
    /// </summary>
    public class Border
    {
        /// <summary>
        /// Instance of Border element
        /// </summary>
        public DocumentFormat.OpenXml.Spreadsheet.Border Element { get; }
        /// <summary>
        /// Sets the left border
        /// </summary>
        public BorderStyleValues Left
        {
            set { Element.LeftBorder = new DocumentFormat.OpenXml.Spreadsheet.LeftBorder { Style = (DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues)value }; }
        }
        /// <summary>
        /// Sets the right border
        /// </summary>
        public BorderStyleValues Right
        {
            set { Element.RightBorder = new DocumentFormat.OpenXml.Spreadsheet.RightBorder { Style = (DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues)value }; }
        }
        /// <summary>
        /// Sets the top border
        /// </summary>
        public BorderStyleValues Top
        {
            set { Element.TopBorder = new DocumentFormat.OpenXml.Spreadsheet.TopBorder { Style = (DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues)value }; }
        }
        /// <summary>
        /// Sets the bottom border
        /// </summary>
        public BorderStyleValues Bottom
        {
            set { Element.BottomBorder = new DocumentFormat.OpenXml.Spreadsheet.BottomBorder { Style = (DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues)value }; }
        }
        /// <summary>
        /// Sets the set border style.
        /// </summary>
        public BorderStyleValues SetBorderStyle
        {
            set
            {
                var style = (DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues)value;
                Element.TopBorder = new DocumentFormat.OpenXml.Spreadsheet.TopBorder { Style = style };
                Element.RightBorder = new DocumentFormat.OpenXml.Spreadsheet.RightBorder { Style = style };
                Element.BottomBorder = new DocumentFormat.OpenXml.Spreadsheet.BottomBorder { Style = style };
                Element.LeftBorder = new DocumentFormat.OpenXml.Spreadsheet.LeftBorder { Style = style };
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Border"/> class.
        /// </summary>
        /// <param name="border">Spreadsheet border</param>
        public Border(DocumentFormat.OpenXml.Spreadsheet.Border border = null)
        {
            Element = border ?? new DocumentFormat.OpenXml.Spreadsheet.Border();
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="Border"/> class.
        /// </summary>
        /// <param name="borderStyle">The border style.</param>
        public Border(BorderStyleValues borderStyle)
            : this()
        {
            SetBorderStyle = borderStyle;
        }

        /// <summary>
        /// Compare content with 'border'
        /// </summary>
        /// <param name="border">Spreadsheet border for compare</param>
        public bool IsContentSame(DocumentFormat.OpenXml.Spreadsheet.Border border)
        {
            return Utils.CompareXml(border.OuterXml, Element.OuterXml);
        }
    }
}