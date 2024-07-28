using DocumentFormat.OpenXml.Spreadsheet;
using OfficeDocuments.Excel.DataClasses;
using OfficeDocuments.Excel.Interfaces;
using Alignment = OfficeDocuments.Excel.Styles.Alignment;
using Border = OfficeDocuments.Excel.Styles.Border;
using Excel_Styles_Alignment = OfficeDocuments.Excel.Styles.Alignment;
using Fill = OfficeDocuments.Excel.Styles.Fill;
using Font = OfficeDocuments.Excel.Styles.Font;
using NumberingFormat = OfficeDocuments.Excel.Styles.NumberingFormat;
using Styles_Alignment = OfficeDocuments.Excel.Styles.Alignment;
using Styles_Border = OfficeDocuments.Excel.Styles.Border;
using Styles_Fill = OfficeDocuments.Excel.Styles.Fill;
using Styles_Font = OfficeDocuments.Excel.Styles.Font;
using Styles_NumberingFormat = OfficeDocuments.Excel.Styles.NumberingFormat;

namespace OfficeDocuments.Excel.Factory.Impl
{
    public class StyleFactory : IStyleFactory
    {
        public IStyle CreateStyle(Stylesheet stylesheet, Styles_Font font = null, Styles_Fill fill = null, Styles_Border border = null,
            Styles_NumberingFormat numberFormat = null, Styles_Alignment alignment = null)
        {
            return new Style(stylesheet, font, fill, border, numberFormat, alignment);
        }

        public IStyle CreateStyle(Stylesheet stylesheet, int fontId = 0, int fillId = 0, int borderId = 0, int numberFormatId = 0,
            Excel_Styles_Alignment alignment = null)
        {
            return new Style(stylesheet, fontId, fillId, borderId, numberFormatId, alignment);
        }

        public IStyle CreateStyle(Stylesheet stylesheet, uint styleIndex)
        {
            return new Style(stylesheet, styleIndex);
        }
    }
}