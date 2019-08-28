using System.Linq;
using OpenXmlApi.Interfaces;
using OpenXmlApi.Styles;

namespace OpenXmlApi
{
    internal abstract class Base : IBase
    {
        public IWorksheet Worksheet { get; protected set; }
        public IStyle Style { get; protected set; }

        protected Base(IWorksheet worksheet, IStyle cellStyle = null)
        {
            this.Worksheet = worksheet;
            AddStyle(cellStyle);
        }
        protected Base(IWorksheet worksheet, uint cellStyle)
        {
            this.Worksheet = worksheet;

            if (cellStyle > 0)
            {
                this.Style = new Style(this.Worksheet.Spreadsheet.Stylesheet, cellStyle);
                AddStyle(this.Style);
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
                this.Style = this.Style?.CreateMergedStyle(style) ?? style;
            }

            return this.Style;
        }
    }
}