using DocumentFormat.OpenXml;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Extensions;
using Color = System.Drawing.Color;

namespace OfficeDocuments.Excel.Styles
{
    /// <summary>
    /// Class of Font
    /// </summary>
    public class Font
    {
        /// <summary>
        /// Instance of Font element
        /// </summary>
        public DocumentFormat.OpenXml.Spreadsheet.Font Element { get; }
        /// <summary>
        /// Sets the bold.
        /// </summary>
        public bool Bold
        {
            set { Element.Bold = new DocumentFormat.OpenXml.Spreadsheet.Bold { Val = value }; }
        }
        /// <summary>
        /// Sets the italic.
        /// </summary>
        public bool Italic
        {
            set { Element.Italic = new DocumentFormat.OpenXml.Spreadsheet.Italic { Val = value }; }
        }
        /// <summary>
        /// Sets the underline.
        /// </summary>
        public UnderlineValues Underline
        {
            set 
            { 
                Element.Underline = new DocumentFormat.OpenXml.Spreadsheet.Underline
                {
                    Val = value switch
                    {
                        UnderlineValues.None => DocumentFormat.OpenXml.Spreadsheet.UnderlineValues.None,
                        UnderlineValues.Single => DocumentFormat.OpenXml.Spreadsheet.UnderlineValues.Single,
                        UnderlineValues.Double => DocumentFormat.OpenXml.Spreadsheet.UnderlineValues.Double,
                        UnderlineValues.SingleAccounting => DocumentFormat.OpenXml.Spreadsheet.UnderlineValues.SingleAccounting,
                        UnderlineValues.DoubleAccounting => DocumentFormat.OpenXml.Spreadsheet.UnderlineValues.DoubleAccounting,
                        _ => DocumentFormat.OpenXml.Spreadsheet.UnderlineValues.None
                    },
                }; 
            }
        }
        /// <summary>
        /// Sets the size of the font.
        /// </summary>
        public double FontSize
        {
            set { Element.FontSize = new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = value }; }
        }
        /// <summary>
        /// Sets the name of the font.
        /// </summary>
        public FontNameValues FontName
        {
            set { Element.FontName = new DocumentFormat.OpenXml.Spreadsheet.FontName { Val = value.ToString() }; }
        }
        /// <summary>
        /// Sets the color.
        /// </summary>
        public Color Color
        {
            set { Element.Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = new HexBinaryValue { Value = Utils.ArgbHexConverter(value) } }; }
        }
        /// <summary>
        /// Sets the color with the ARGB value.
        /// </summary>
        public string ArgbHexColor
        {
            set { Element.Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = new HexBinaryValue { Value = value } }; }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Font"/> class.
        /// </summary>
        /// <param name="font">Spreadsheet font.</param>
        public Font(DocumentFormat.OpenXml.Spreadsheet.Font font = null)
        {
            Element = font ?? new DocumentFormat.OpenXml.Spreadsheet.Font();
        }

        /// <summary>
        /// Compare content with 'font'
        /// </summary>
        /// <param name="font">Spreadsheet font for compare</param>
        public bool IsContentSame(DocumentFormat.OpenXml.Spreadsheet.Font font)
        {
            return font.OuterXml.CompareXml(Element.OuterXml);
        }
    }
}