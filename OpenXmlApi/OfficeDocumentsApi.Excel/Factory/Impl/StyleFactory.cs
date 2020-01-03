using DocumentFormat.OpenXml.Spreadsheet;
using OfficeDocumentsApi.Excel.DataClasses;
using OfficeDocumentsApi.Excel.Interfaces;
using Alignment = OfficeDocumentsApi.Excel.Styles.Alignment;
using Border = OfficeDocumentsApi.Excel.Styles.Border;
using Fill = OfficeDocumentsApi.Excel.Styles.Fill;
using Font = OfficeDocumentsApi.Excel.Styles.Font;
using NumberingFormat = OfficeDocumentsApi.Excel.Styles.NumberingFormat;

namespace OfficeDocumentsApi.Excel.Factory.Impl
{
    public class StyleFactory : IStyleFactory
    {
        public IStyle CreateStyle(Stylesheet stylesheet, Font font = null, Fill fill = null, Border border = null,
            NumberingFormat numberFormat = null, Alignment alignment = null)
        {
            return new Style(stylesheet, font, fill, border, numberFormat, alignment);
        }

        public IStyle CreateStyle(Stylesheet stylesheet, int fontId = 0, int fillId = 0, int borderId = 0, int numberFormatId = 0,
            Alignment alignment = null)
        {
            return new Style(stylesheet, fontId, fillId, borderId, numberFormatId, alignment);
        }

        public IStyle CreateStyle(Stylesheet stylesheet, uint styleIndex)
        {
            return new Style(stylesheet, styleIndex);
        }
    }
}