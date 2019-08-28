using OpenXmlApi.Styles;
using OpenXmlApi.Emums;
using Color = System.Drawing.Color;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenXmlApi.Test
{
    [TestClass]
    public class StyleTest
    {
        [TestMethod]
        public void BasicStyle()
        {
            var filePath = GetFilepath("doc1.xlsx");
            using (var w = Spreadsheet.Create(filePath))
            {
                var s = w.CreateStyle();

                Assert.IsNotNull(s.Element, "OpenXml element for style in not set");
                Assert.IsTrue(s.FontId == 0, "Font is not default");
                Assert.IsTrue(s.FillId == 0, "Fill is not default");
                Assert.IsTrue(s.NumberFormatId == 0, "NumberFormat is not default");
                Assert.IsTrue(s.BorderId == 0, "Border is not default");
                Assert.IsTrue(s.StyleIndex == 0, $"Its not required to create new style: {s.StyleIndex}");
            }
        }

        [TestMethod]
        public void SpecificFontStyle()
        {
            var filePath = GetFilepath("doc2.xlsx");
            using (var w = Spreadsheet.Create(filePath))
            {
                var s = w.CreateStyle(
                    new Font { FontSize = 15, Color = Color.Blue, FontName = FontNameValues.Tahoma, Bold = true, Italic = true, Underline = UnderlineValues.Double }
                );

                Assert.IsNotNull(s.Element, "OpenXml element for style in not set");
                Assert.IsTrue(s.FontId > 0, "Font is default");
                Assert.IsTrue(s.FillId == 0, "Fill is not default");
                Assert.IsTrue(s.NumberFormatId == 0, "NumberFormat is not default");
                Assert.IsTrue(s.BorderId == 0, "Border is not default");
                Assert.IsTrue(s.StyleIndex > 0, $"Style is default. Index: {s.StyleIndex}");
            }
        }

        [TestMethod]
        public void SpecificFillStyle()
        {
            var filePath = GetFilepath("doc3.xlsx");
            using (var w = Spreadsheet.Create(filePath))
            {
                var s = w.CreateStyle(
                    fill: new Fill(Color.Blue, Color.White)
                );

                Assert.IsNotNull(s.Element, "OpenXml element for style in not set");
                Assert.IsTrue(s.FontId == 0, "Font is not default");
                Assert.IsTrue(s.FillId > 0, "Fill is default");
                Assert.IsTrue(s.NumberFormatId == 0, "NumberFormat is not default");
                Assert.IsTrue(s.BorderId == 0, "Border is not default");
                Assert.IsTrue(s.StyleIndex > 0, $"Style is default. Index: {s.StyleIndex}");
            }
        }

        [TestMethod]
        public void SpecificBorderStyle()
        {
            var filePath = GetFilepath("doc4.xlsx");
            using (var w = Spreadsheet.Create(filePath))
            {
                var b = new Border
                {
                    Top = BorderStyleValues.Double,
                    Right = BorderStyleValues.Double,
                    Bottom = BorderStyleValues.Double,
                    Left = BorderStyleValues.Double
                };

                var s = w.CreateStyle(
                    border: b
                );

                Assert.IsNotNull(s.Element, "OpenXml element for style in not set");
                Assert.IsTrue(s.FontId == 0, "Font is not default");
                Assert.IsTrue(s.FillId == 0, "Fill is not default");
                Assert.IsTrue(s.NumberFormatId == 0, "NumberFormat is not default");
                Assert.IsTrue(s.BorderId > 0, "Border is default");
                Assert.IsTrue(s.StyleIndex > 0, $"Style is default. Index: {s.StyleIndex}");
            }
        }

        [TestMethod]
        public void SpecificBorderStyle1()
        {
            var filePath = GetFilepath("doc5.xlsx");
            using (var w = Spreadsheet.Create(filePath))
            {
                var s = w.CreateStyle(
                    border: new Border(BorderStyleValues.Medium)
                );

                Assert.IsNotNull(s.Element, "OpenXml element for style in not set");
                Assert.IsTrue(s.FontId == 0, "Font is not default");
                Assert.IsTrue(s.FillId == 0, "Fill is not default");
                Assert.IsTrue(s.NumberFormatId == 0, "NumberFormat is not default");
                Assert.IsTrue(s.BorderId > 0, "Border is default");
                Assert.IsTrue(s.StyleIndex > 0, $"Style is default. Index: {s.StyleIndex}");
            }
        }

        [TestMethod]
        public void SpecificNumberFormatStyle()
        {
            var filePath = GetFilepath("doc6.xlsx");
            using (var w = Spreadsheet.Create(filePath))
            {
                var s = w.CreateStyle(
                    numberFormat: new NumberingFormat("@")
                );

                Assert.IsNotNull(s.Element, "OpenXml element for style in not set");
                Assert.IsTrue(s.FontId == 0, "Font is not default");
                Assert.IsTrue(s.FillId == 0, "Fill is not default");
                Assert.IsTrue(s.NumberFormatId > 0, "NumberFormat is not default");
                Assert.IsTrue(s.BorderId == 0, "Border is default");
                Assert.IsTrue(s.StyleIndex > 0, $"Style is default. Index: {s.StyleIndex}");
            }
        }

        [TestMethod]
        public void SpecificAlignmentStyle()
        {
            var filePath = GetFilepath("doc7.xlsx");
            using (var w = Spreadsheet.Create(filePath))
            {
                var s = w.CreateStyle(
                    alignment: new Alignment
                    {
                        Horizontal = HorizontalAlignmentValues.Center,
                        Vertical = VerticalAlignmentValues.Center,
                        JustifyLastLine = true,
                        ShrinkToFit = true,
                        WrapText = true
                    }
                );

                Assert.IsNotNull(s.Element, "OpenXml element for style in not set");
                Assert.IsTrue(s.FontId == 0, "Font is not default");
                Assert.IsTrue(s.FillId == 0, "Fill is not default");
                Assert.IsTrue(s.NumberFormatId == 0, "NumberFormat is not default");
                Assert.IsTrue(s.BorderId == 0, "Border is default");
                Assert.IsTrue(s.Element.Alignment != null, "Alignment is not set");
                Assert.IsTrue(s.StyleIndex > 0, $"Style is default. Index: {s.StyleIndex}");
            }
        }

        [TestMethod]
        public void MergeStyles()
        {
            var filePath = GetFilepath("doc8.xlsx");
            using (var w = Spreadsheet.Create(filePath))
            {
                var s1 = w.CreateStyle(
                    font: new Font { FontSize = 15, Color = Color.Brown, FontName = FontNameValues.Calibri },
                    border: new Border(BorderStyleValues.Double)
                );
                var s2 = w.CreateStyle(
                    font: new Font { FontSize = 20, Color = Color.Brown, FontName = FontNameValues.Tahoma },
                    numberFormat: new NumberingFormat("0x")
                );

                var s = s1.CreateMergedStyle(s2);

                Assert.IsNotNull(s.Element, "OpenXml element for style in not set");
                Assert.IsTrue(s.FontId > 0 && s.FontId == s2.FontId, $"Font is '{s.FontId}', it should be '{s2.FontId}'");
                Assert.IsTrue(s.FillId == 0, "Fill is not default");
                Assert.IsTrue(s.NumberFormatId > 0 && s.NumberFormatId == s2.NumberFormatId, $"NumberFormat is '{s.NumberFormatId}', it should be '{s2.NumberFormatId}'");
                Assert.IsTrue(s.BorderId > 0 && s.BorderId == s1.BorderId, $"Border is '{s.BorderId}', it should be '{s1.BorderId}'");
                Assert.IsTrue(s.StyleIndex > 0, $"Style is default. Index: {s.StyleIndex}");
            }
        }

        [TestMethod]
        public void MergeStylesToKnownStyle()
        {
            var filePath = GetFilepath("doc9.xlsx");
            using (var w = Spreadsheet.Create(filePath))
            {
                var s_old = w.CreateStyle(
                    font: new Font { FontSize = 20, Color = Color.Brown, FontName = FontNameValues.Tahoma },
                    border: new Border(BorderStyleValues.Double),
                    numberFormat: new NumberingFormat("0x")
                );

                var s1 = w.CreateStyle(
                    font: new Font { FontSize = 15, Color = Color.Brown, FontName = FontNameValues.Calibri },
                    border: new Border(BorderStyleValues.Double)
                );
                var s2 = w.CreateStyle(
                    font: new Font { FontSize = 20, Color = Color.Brown, FontName = FontNameValues.Tahoma },
                    numberFormat: new NumberingFormat("0x")
                );

                var s = s1.CreateMergedStyle(s2);

                Assert.IsNotNull(s.Element, "OpenXml element for style in not set");
                Assert.IsTrue(s.FontId == s_old.FontId, $"Font is '{s.FontId}', it should be '{s_old.FontId}'");
                Assert.IsTrue(s.FillId == s_old.FillId, $"Fill is '{s.FillId}', it should be '{s_old.FillId}'");
                Assert.IsTrue(s.NumberFormatId == s_old.NumberFormatId, $"NumberFormat is '{s.NumberFormatId}', it should be '{s_old.NumberFormatId}'");
                Assert.IsTrue(s.BorderId == s_old.BorderId, $"Border is '{s.BorderId}', it should be '{s_old.BorderId}'");
                Assert.IsTrue(s.Element.Alignment == null, "Alignment on 's' should not be set");
                Assert.IsTrue(s_old.Element.Alignment == null, "Alignment on 's_old' should not be set");
                Assert.IsTrue(s.StyleIndex == s_old.StyleIndex, $"Style is '{s.StyleIndex}', it should be '{s_old.StyleIndex}'");
            }
        }

        [TestMethod]
        public void MergeStylesWithNull()
        {
            var filePath = GetFilepath("doc10.xlsx");
            using (var w = Spreadsheet.Create(filePath))
            {
                var s_old = w.CreateStyle(
                    font: new Font { FontSize = 20, Color = Color.Brown, FontName = FontNameValues.Tahoma, Bold = true },
                    border: new Border(BorderStyleValues.Double),
                    numberFormat: new NumberingFormat("0x")
                );

                var s = s_old.CreateMergedStyle(null);

                Assert.IsNotNull(s.Element, "OpenXml element for style in not set");
                Assert.IsTrue(s.FontId == s_old.FontId, $"Font is '{s.FontId}', it should be '{s_old.FontId}'");
                Assert.IsTrue(s.FillId == s_old.FillId, $"Fill is '{s.FillId}', it should be '{s_old.FillId}'");
                Assert.IsTrue(s.NumberFormatId == s_old.NumberFormatId, $"NumberFormat is '{s.NumberFormatId}', it should be '{s_old.NumberFormatId}'");
                Assert.IsTrue(s.BorderId == s_old.BorderId, $"Border is '{s.BorderId}', it should be '{s_old.BorderId}'");
                Assert.IsTrue(s.Element.Alignment == null, "Alignment on 's' should not be set");
                Assert.IsTrue(s_old.Element.Alignment == null, "Alignment on 's_old' should not be set");
                Assert.IsTrue(s.StyleIndex == s_old.StyleIndex, $"Style is '{s.StyleIndex}', it should be '{s_old.StyleIndex}'");
            }
        }

        private string GetFilepath(string filename)
        {
            return TestSettings.GetFilepath(this, filename);
        }
    }
}