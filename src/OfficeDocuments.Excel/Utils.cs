﻿using System.Linq;
using System.Xml.Linq;
using OfficeDocuments.Excel.Styles;

namespace OfficeDocuments.Excel
{
    /// <summary>
    /// Class of utilities
    /// </summary>
    public static class Utils
    {
        /// <summary>
        /// Convert color from 'System.Drawing.Color' to argb hex representation
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        public static string ArgbHexConverter(System.Drawing.Color c)
        {
            return $"{c.A:X2}{c.R:X2}{c.G:X2}{c.B:X2}";
        }

        /// <summary>
        /// Create new font by merging two fonts
        /// </summary>
        public static Font MergeFonts(DocumentFormat.OpenXml.Spreadsheet.Font font1, DocumentFormat.OpenXml.Spreadsheet.Font font2)
        {
            var a = XDocument.Parse(font1.OuterXml);
            var b = XDocument.Parse(font2.OuterXml);

            var element = new DocumentFormat.OpenXml.Spreadsheet.Font(a.MergeXml(b).ToString());

            return new Font(element);
        }

        /// <summary>
        /// Create new fill by merging two fills
        /// </summary>
        public static Fill MergeFills(DocumentFormat.OpenXml.Spreadsheet.Fill fill1, DocumentFormat.OpenXml.Spreadsheet.Fill fill2)
        {
            var a = XDocument.Parse(fill1.OuterXml);
            var b = XDocument.Parse(fill2.OuterXml);

            var element = new DocumentFormat.OpenXml.Spreadsheet.Fill(a.MergeXml(b).ToString());

            return new Fill(element);
        }

        /// <summary>
        /// Create new border by merging two borders
        /// </summary>
        public static Border MergeBorders(DocumentFormat.OpenXml.Spreadsheet.Border border1, DocumentFormat.OpenXml.Spreadsheet.Border border2)
        {
            var a = XDocument.Parse(border1.OuterXml);
            var b = XDocument.Parse(border2.OuterXml);

            var element = new DocumentFormat.OpenXml.Spreadsheet.Border(a.MergeXml(b).ToString());

            return new Border(element);
        }

        private static XDocument MergeXml(this XDocument xd1, XDocument xd2)
        {
            var docs = new XDocument(
                new XElement(xd2.Root.Name,
                    xd2.Root.Attributes()
                        .Concat(xd1.Root.Attributes())
                        .GroupBy(g => g.Name)
                        .Select(s => s.First()),
                    xd2.Root.Elements()
                        .Concat(xd1.Root.Elements())
                        .GroupBy(g => g.Name)
                        .Select(s => s.First())
                ));

            return docs;
        }
    }
}