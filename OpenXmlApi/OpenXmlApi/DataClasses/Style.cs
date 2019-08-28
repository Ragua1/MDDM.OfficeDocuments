using System;
using System.Linq;
using OpenXmlApi.Interfaces;

namespace OpenXmlApi.Styles
{
    internal class Style : IStyle
    {
        public DocumentFormat.OpenXml.Spreadsheet.Stylesheet Stylesheet { get; }
        public DocumentFormat.OpenXml.Spreadsheet.CellFormat Element { get; }
        public uint StyleIndex { get; }

        public int FontId => Convert.ToInt32(this.Element.FontId.Value);
        public int FillId => Convert.ToInt32(this.Element.FillId.Value);
        public int BorderId => Convert.ToInt32(this.Element.BorderId.Value);
        public int NumberFormatId => Convert.ToInt32(this.Element.NumberFormatId?.Value ?? 0);

        internal Style(DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet, Font font = null, Fill fill = null, Border border = null, NumberingFormat numberFormat = null)
            : this(stylesheet, GetFontId(stylesheet, font), GetFillId(stylesheet, fill), GetBorderId(stylesheet, border), GetNumberFormatId(stylesheet, numberFormat))
        { }
        internal Style(DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet, Font font = null, Fill fill = null, Border border = null, NumberingFormat numberFormat = null, Alignment alignment = null)
            : this(stylesheet, GetFontId(stylesheet, font), GetFillId(stylesheet, fill), GetBorderId(stylesheet, border), GetNumberFormatId(stylesheet, numberFormat), alignment)
        { }
        internal Style(DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet, int fontId = 0, int fillId = 0, int borderId = 0, int numberFormatId = 0, Alignment alignment = null)
        {
            this.Stylesheet = stylesheet;
            this.Element = new DocumentFormat.OpenXml.Spreadsheet.CellFormat
            {
                FormatId = Convert.ToUInt32(0),
                FontId = Convert.ToUInt32(fontId),
                FillId = Convert.ToUInt32(fillId),
                BorderId = Convert.ToUInt32(borderId)
            };

            if (numberFormatId >= 0)
            {
                this.Element.NumberFormatId = Convert.ToUInt32(numberFormatId);
            }

            if (alignment != null)
            {
                this.Element.Alignment = (DocumentFormat.OpenXml.Spreadsheet.Alignment)alignment.Element.CloneNode(true);
            }

            this.StyleIndex = GetStyleIndex(stylesheet, this.Element);
        }
        internal Style(DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet, uint styleIndex)
        {
            this.Stylesheet = stylesheet;

            var cfs = stylesheet.CellFormats ?? new DocumentFormat.OpenXml.Spreadsheet.CellFormats();
            var cellFormats = cfs.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ToList();
            this.Element = cellFormats.ElementAt(Convert.ToInt32(styleIndex));

            this.StyleIndex = styleIndex;
        }

        public IStyle CreateMergedStyle(IStyle style)
        {
            int fontId = this.FontId, fillId = this.FillId, borderId = this.BorderId, numberFormatId = this.NumberFormatId;
            var alignment = this.Element.Alignment != null ? new Alignment(this.Element.Alignment) : null;
            if (style == null)
            {
                return this;// new Style(this.Stylesheet, fontId, fillId, borderId, numberFormatId, alignment);
            }

            if (fontId != style.FontId && style.FontId > 0) // Id == 0 is default style
            {
                var fonts = this.Stylesheet.Fonts.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ToList();
                var font1 = fonts.ElementAt(this.FontId);
                var font2 = fonts.ElementAt(style.FontId);
                var font = Utils.MergeFonts(font1, font2);
                fontId = GetFontId(style.Stylesheet, font);
            }

            if (fillId != style.FillId && style.FillId > 0) // Id == 0 is default style
            {
                var fills = this.Stylesheet.Fills.Elements<DocumentFormat.OpenXml.Spreadsheet.Fill>().ToList();
                var fill1 = fills.ElementAt(this.FillId);
                var fill2 = fills.ElementAt(style.FillId);
                var fill = Utils.MergeFills(fill1, fill2);
                fillId = GetFillId(style.Stylesheet, fill);
            }

            if (borderId != style.BorderId && style.BorderId > 0) // Id == 0 is default style
            {
                var borders = this.Stylesheet.Borders.Elements<DocumentFormat.OpenXml.Spreadsheet.Border>().ToList();
                var border1 = borders.ElementAt(this.BorderId);
                var border2 = borders.ElementAt(style.BorderId);
                var border = Utils.MergeBorders(border1, border2);
                borderId = GetBorderId(style.Stylesheet, border);
            }

            if (numberFormatId != style.NumberFormatId && style.NumberFormatId > 0) // Id == 0 is default style
            {
                numberFormatId = style.NumberFormatId; // Alignment cannot be merged
            }

            if (!string.IsNullOrEmpty(style.Element.Alignment?.InnerXml))
            {
                if (string.IsNullOrEmpty(this.Element.Alignment?.InnerXml))
                {
                    alignment = new Alignment(style.Element.Alignment); // Alignment cannot be merged
                }
            }

            return new Style(this.Stylesheet, fontId, fillId, borderId, numberFormatId, alignment);
        }

        private static int GetFontId(DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet, Font font)
        {
            var fontId = 0;
            if (font?.Element != null)
            {
                var fonts = stylesheet.Fonts ?? (stylesheet.Fonts = new DocumentFormat.OpenXml.Spreadsheet.Fonts());
                var elms = fonts.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ToList();

                fontId = elms.FindIndex(font.IsContentSame);

                if (fontId <= 0) // Id == 0 is default style, Id < 0 element not exist yet
                {
                    fonts.Append(font.Element);
                    fontId = elms.Count;
                }
            }
            return fontId;
        }

        private static int GetFillId(DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet, Fill fill)
        {
            var fillId = 0;
            if (fill?.Element != null)
            {
                var fills = stylesheet.Fills ?? (stylesheet.Fills = new DocumentFormat.OpenXml.Spreadsheet.Fills());
                var elms = fills.Elements<DocumentFormat.OpenXml.Spreadsheet.Fill>().ToList();

                fillId = elms.FindIndex(fill.IsContentSame);

                if (fillId <= 0) // Id == 0 is default style, Id < 0 element not exist yet
                {
                    fills.Append(fill.Element);
                    fillId = elms.Count;
                }
            }
            return fillId;
        }

        private static int GetBorderId(DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet, Border border)
        {
            var borderId = 0;
            if (border?.Element != null)
            {
                var borders = stylesheet.Borders ?? (stylesheet.Borders = new DocumentFormat.OpenXml.Spreadsheet.Borders());
                var elms = borders.Elements<DocumentFormat.OpenXml.Spreadsheet.Border>().ToList();

                borderId = elms.FindIndex(border.IsContentSame);

                if (borderId <= 0) // Id == 0 is default style, Id < 0 element not exist yet
                {
                    borders.Append(border.Element);
                    borderId = elms.Count;
                }
            }
            return borderId;
        }

        private static int GetNumberFormatId(DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet, NumberingFormat numberFormat)
        {
            var numberFormatId = 0;
            if (numberFormat?.Element != null)
            {
                var numberingFormats = stylesheet.NumberingFormats ?? (stylesheet.NumberingFormats = new DocumentFormat.OpenXml.Spreadsheet.NumberingFormats());
                var elms = numberingFormats.Elements<DocumentFormat.OpenXml.Spreadsheet.NumberingFormat>().ToList();

                var numFormat = elms.FirstOrDefault(numberFormat.IsContentSame);

                if (numFormat == null)
                {
                    numberingFormats.Append(numberFormat.Element);
                    numberFormatId = Convert.ToInt32(numberFormat.Element.NumberFormatId.Value);
                }
                else
                {
                    numberFormatId = Convert.ToInt32(numFormat.NumberFormatId.Value);
                }
            }
            return numberFormatId;
        }

        private static uint GetStyleIndex(DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet, DocumentFormat.OpenXml.Spreadsheet.CellFormat element)
        {
            var cfs = stylesheet.CellFormats ?? new DocumentFormat.OpenXml.Spreadsheet.CellFormats();
            var cellFormats = cfs.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ToList();
            if (cellFormats.Any())
            {
                for (uint i = 0; i < cellFormats.Count; i++)
                {
                    var elm = cellFormats[Convert.ToInt32(i)];
                    if (Equals(element, elm))
                    {
                        return i;
                    }
                }
            }

            cfs.Append(element);
            cfs.Count = Convert.ToUInt32(cfs.Count());
            return (uint)(cfs.Count() - 1);
        }

        private static bool Equals(DocumentFormat.OpenXml.Spreadsheet.CellFormat style1, DocumentFormat.OpenXml.Spreadsheet.CellFormat style2)
        {
            var res = style1.FontId.Value == style2.FontId.Value
                      && style1.FillId.Value == style2.FillId.Value
                      && style1.BorderId.Value == style2.BorderId.Value;

            if (style1.NumberFormatId.HasValue == style2.NumberFormatId.HasValue)
            {
                if (style1.NumberFormatId.HasValue)
                {
                    res &= style1.NumberFormatId.Value == style2.NumberFormatId.Value;
                }
            }
            else
            {
                return false;
            }

            if (style1.Alignment != null && style2.Alignment != null)
            {
                res &= Utils.CompareXml(style1.Alignment.OuterXml, style2.Alignment.OuterXml);
            }
            else
            {
                if (style1.Alignment == null && style2.Alignment != null
                    || style1.Alignment != null && style2.Alignment == null)
                {
                    return false;
                }
            }


            return res;
        }
    }
}