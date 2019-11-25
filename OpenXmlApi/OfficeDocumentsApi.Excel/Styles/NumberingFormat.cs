using System.Collections.Generic;
using DocumentFormat.OpenXml;

namespace OfficeDocumentsApi.Excel.Styles
{
    /// <summary>
    /// Class of NumberingFormat
    /// </summary>
    public class NumberingFormat
    {
        /// <summary>
        /// Instance of NumberingFormat element
        /// </summary>
        public DocumentFormat.OpenXml.Spreadsheet.NumberingFormat Element { get; }
        private static readonly Dictionary<string, uint> DefaultNumberFormats = new Dictionary<string, uint>
        {
            { "General", 0 },
            { "0", 1 },
            { "0.00", 2 },
            { "#,##0", 3 },
            { "#,##0.00", 4 },
            { "0%", 9 },
            { "0.00%", 10 },
            { "0.00E+00", 11 },
            { "# ?/?", 12 },
            { "# ??/??", 13 },
            { "d/m/yyyy", 14 },
            { "d-mmm-yy", 15 },
            { "d-mmm", 16 },
            { "mmm-yy", 17 },
            { "h:mm tt", 18 },
            { "h:mm:ss tt", 19 },
            { "H:mm", 20 },
            { "H:mm:ss", 21 },
            { "m/d/yyyy H:mm", 22 },
            { "#,##0 ;(#,##0)", 37 },
            { "#,##0 ;[Red](#,##0)", 38 },
            { "#,##0.00;(#,##0.00)", 39 },
            { "#,##0.00;[Red](#,##0.00)", 40 },
            { "mm:ss", 45 },
            { "[h]:mm:ss", 46 },
            { "mmss.0", 47 },
            { "##0.0E+0", 48 },
            { "@", 49 },
        };
        private static uint excelIndex = 170;

        /// <summary>
        /// Initializes a new instance of the <see cref="NumberingFormat"/> class.
        /// </summary>
        /// <param name="formatCode">The format code.</param>
        public NumberingFormat(string formatCode)
            : this(string.IsNullOrEmpty(formatCode) ? null : new DocumentFormat.OpenXml.Spreadsheet.NumberingFormat { FormatCode = formatCode })
        { }
        /// <summary>
        /// Initializes a new instance of the <see cref="NumberingFormat"/> class.
        /// </summary>
        /// <param name="numberFormat">Spreadsheet number format.</param>
        public NumberingFormat(DocumentFormat.OpenXml.Spreadsheet.NumberingFormat numberFormat = null)
        {
            Element = numberFormat ?? new DocumentFormat.OpenXml.Spreadsheet.NumberingFormat();

            if (string.IsNullOrEmpty(Element.FormatCode.Value))
            {
                Element.FormatCode = StringValue.FromString("General");
            }

            if (DefaultNumberFormats.ContainsKey(Element.FormatCode.Value))
            {
                Element.NumberFormatId = DefaultNumberFormats[Element.FormatCode.Value];
            }
            else
            {
                DefaultNumberFormats.Add(Element.FormatCode.Value, excelIndex);
                Element.NumberFormatId = excelIndex++;
            }
        }

        /// <summary>
        /// Compare content with 'numberFormat'
        /// </summary>
        /// <param name="numberFormat">Spreadsheet number format for compare</param>
        public bool IsContentSame(DocumentFormat.OpenXml.Spreadsheet.NumberingFormat numberFormat)
        {
            return Utils.CompareXml(numberFormat.OuterXml, Element.OuterXml);
        }
    }
}