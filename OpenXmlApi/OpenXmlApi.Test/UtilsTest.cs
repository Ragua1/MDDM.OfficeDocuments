using Microsoft.VisualStudio.TestTools.UnitTesting;
using Color = System.Drawing.Color;

namespace OpenXmlApi.Test
{
    [TestClass]
    public class UtilsTest
    {
        [TestMethod]
        public void ColorConverter()
        {
            var color = Color.Blue;
            var argbHex = Utils.ArgbHexConverter(color);
            Assert.AreEqual(argbHex, $"{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}");
        }

        [TestMethod]
        public void FontMerge()
        {
            var font1 = new DocumentFormat.OpenXml.Spreadsheet.Font
            {
                Bold = new DocumentFormat.OpenXml.Spreadsheet.Bold { Val = true },
                Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = Utils.ArgbHexConverter(Color.Black) },
            };
            var font2 = new DocumentFormat.OpenXml.Spreadsheet.Font
            {
                Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = Utils.ArgbHexConverter(Color.Red) },
                FontSize = new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = 12 }
            };
            var font_res = new DocumentFormat.OpenXml.Spreadsheet.Font
            {
                Bold = new DocumentFormat.OpenXml.Spreadsheet.Bold { Val = true },
                Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = Utils.ArgbHexConverter(Color.Red) },
                FontSize = new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = 12 }
            };

            var font_merged = Utils.MergeFonts(font1, font2);

            Assert.IsTrue(Utils.CompareXml(font_merged.Element.OuterXml, font_res.OuterXml));
        }

        [TestMethod]
        public void FillMerge()
        {
            var fill1 = new DocumentFormat.OpenXml.Spreadsheet.Fill()
            {
                PatternFill = new DocumentFormat.OpenXml.Spreadsheet.PatternFill
                {
                    PatternType = DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid,
                    ForegroundColor = new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor { Rgb = Utils.ArgbHexConverter(Color.Blue) }
                }
            };
            var fill2 = new DocumentFormat.OpenXml.Spreadsheet.Fill()
            {
                PatternFill = new DocumentFormat.OpenXml.Spreadsheet.PatternFill
                {
                    ForegroundColor = new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor { Rgb = Utils.ArgbHexConverter(Color.Green) }
                }
            };
            var fill_res = new DocumentFormat.OpenXml.Spreadsheet.Fill()
            {
                PatternFill = new DocumentFormat.OpenXml.Spreadsheet.PatternFill
                {
                    PatternType = DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid,
                    ForegroundColor = new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor { Rgb = Utils.ArgbHexConverter(Color.Green) }
                }
            };

            var fill_merged = Utils.MergeFills(fill1, fill2);

            Assert.IsTrue(Utils.CompareXml(fill_merged.Element.OuterXml, fill_res.OuterXml));
        }

        [TestMethod]
        public void BorderMerge()
        {
            var border1 = new DocumentFormat.OpenXml.Spreadsheet.Border()
            {
                LeftBorder = new DocumentFormat.OpenXml.Spreadsheet.LeftBorder { Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = Utils.ArgbHexConverter(Color.Aqua) } },
                TopBorder = new DocumentFormat.OpenXml.Spreadsheet.TopBorder { Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = Utils.ArgbHexConverter(Color.YellowGreen) } }
            };
            var Border2 = new DocumentFormat.OpenXml.Spreadsheet.Border()
            {
                RightBorder = new DocumentFormat.OpenXml.Spreadsheet.RightBorder { Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = Utils.ArgbHexConverter(Color.Red) } },
                BottomBorder = new DocumentFormat.OpenXml.Spreadsheet.BottomBorder { Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = Utils.ArgbHexConverter(Color.Green) } }
            };
            var Border_res = new DocumentFormat.OpenXml.Spreadsheet.Border()
            {
                LeftBorder = new DocumentFormat.OpenXml.Spreadsheet.LeftBorder { Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = Utils.ArgbHexConverter(Color.Aqua) } },
                TopBorder = new DocumentFormat.OpenXml.Spreadsheet.TopBorder { Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = Utils.ArgbHexConverter(Color.YellowGreen) } },
                RightBorder = new DocumentFormat.OpenXml.Spreadsheet.RightBorder { Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = Utils.ArgbHexConverter(Color.Red) } },
                BottomBorder = new DocumentFormat.OpenXml.Spreadsheet.BottomBorder { Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = Utils.ArgbHexConverter(Color.Green) } }
            };

            var Border_merged = Utils.MergeBorders(border1, Border2);

            Assert.IsTrue(Utils.CompareXml(Border_merged.Element.OuterXml, Border_res.OuterXml));
        }
    }
}
