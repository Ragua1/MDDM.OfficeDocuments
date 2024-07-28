using System.Linq;
using OfficeDocuments.Excel.Interfaces;

namespace OfficeDocuments.Excel.DataClasses
{
    internal abstract class Base : IBase
    {
        public IWorksheet Worksheet { get; protected set; }
        public IStyle? Style { get; protected set; }

        protected Base(IWorksheet worksheet, IStyle? cellStyle = null)
        {
            Worksheet = worksheet;
            AddStyle(cellStyle);
        }
        protected Base(IWorksheet worksheet, uint cellStyle)
        {
            Worksheet = worksheet;

            if (cellStyle > 0)
            {
                Style = new Style(Worksheet.Spreadsheet.Stylesheet, cellStyle);
                AddStyle(Style);
            }
        }

        private IStyle AddStyle(IStyle style)
        {
            return AddStyle(style, null);
        }

        public virtual IStyle AddStyle(params IStyle[] styles)
        {
            foreach (var style in styles.Where(s => s != null))
            {
                Style = Style?.CreateMergedStyle(style) ?? style;
            }

            return Style;
        }
    }
}