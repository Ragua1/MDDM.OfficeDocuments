using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlApi.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenXmlApi.DataClasses
{
    public class Style : IStyle
    {
        public Stylesheet Stylesheet => throw new NotImplementedException();

        public CellFormat Element => throw new NotImplementedException();

        public uint StyleIndex => throw new NotImplementedException();

        public int FontId => throw new NotImplementedException();

        public int FillId => throw new NotImplementedException();

        public int BorderId => throw new NotImplementedException();

        public int NumberFormatId => throw new NotImplementedException();

        public IStyle CreateMergedStyle(IStyle style)
        {
            throw new NotImplementedException();
        }
    }
}
