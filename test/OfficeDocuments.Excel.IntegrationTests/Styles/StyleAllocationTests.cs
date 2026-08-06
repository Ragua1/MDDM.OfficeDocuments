using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Styles;
using OfficeDocuments.Excel.TestKit;
using Color = System.Drawing.Color;

namespace OfficeDocuments.Excel.IntegrationTests;

public class StyleAllocationTests : SpreadsheetTestBase
{
    [Fact]
    public void BasicStyle()
    {
        using var w = CreateInMemorySpreadsheet();
        var s = w.CreateStyle();

        Assert.Equal(0, s.FontId);
        Assert.Equal(0, s.FillId);
        Assert.Equal(0, s.NumberFormatId);
        Assert.Equal(0, s.BorderId);
        Assert.Equal(0U, s.StyleIndex);
    }

    [Fact]
    public void SpecificFontStyle()
    {
        using var w = CreateInMemorySpreadsheet();
        var s = w.CreateStyle(
            new Font { FontSize = 15, Color = Color.Blue, FontName = FontNameValues.Tahoma, Bold = true, Italic = true, Underline = UnderlineValues.Double }
        );

        Assert.True(s.FontId > 0);
        Assert.Equal(0, s.FillId);
        Assert.Equal(0, s.NumberFormatId);
        Assert.Equal(0, s.BorderId);
        Assert.True(s.StyleIndex > 0);
    }

    [Fact]
    public void SpecificFillStyle()
    {
        using var w = CreateInMemorySpreadsheet();
        var s = w.CreateStyle(
            fill: new Fill(Color.Blue, Color.White)
        );

        Assert.Equal(0, s.FontId);
        Assert.True(s.FillId > 0);
        Assert.Equal(0, s.NumberFormatId);
        Assert.Equal(0, s.BorderId);
        Assert.True(s.StyleIndex > 0);
    }

    [Fact]
    public void SpecificBorderStyle()
    {
        using var w = CreateInMemorySpreadsheet();
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

        Assert.Equal(0, s.FontId);
        Assert.Equal(0, s.FillId);
        Assert.Equal(0, s.NumberFormatId);
        Assert.True(s.BorderId > 0);
        Assert.True(s.StyleIndex > 0);
    }

    [Fact]
    public void SpecificBorderStyle1()
    {
        using var w = CreateInMemorySpreadsheet();
        var s = w.CreateStyle(
            border: new Border(BorderStyleValues.Medium)
        );

        Assert.Equal(0, s.FontId);
        Assert.Equal(0, s.FillId);
        Assert.Equal(0, s.NumberFormatId);
        Assert.True(s.BorderId > 0);
        Assert.True(s.StyleIndex > 0);
    }

    [Fact]
    public void SpecificNumberFormatStyle()
    {
        using var w = CreateInMemorySpreadsheet();
        var s = w.CreateStyle(
            numberFormat: new NumberingFormat("@")
        );

        Assert.Equal(0, s.FontId);
        Assert.Equal(0, s.FillId);
        Assert.Equal(49, s.NumberFormatId);
        Assert.Equal(0, s.BorderId);
        Assert.True(s.StyleIndex > 0);
    }

    [Fact]
    public void SpecificAlignmentStyle()
    {
        using var w = CreateInMemorySpreadsheet();
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

        Assert.Equal(0, s.FontId);
        Assert.Equal(0, s.FillId);
        Assert.Equal(0, s.NumberFormatId);
        Assert.Equal(0, s.BorderId);
        Assert.NotNull(StylesheetProbe.Alignment(s));
        Assert.True(s.StyleIndex > 0);
    }

    [Fact]
    public void MergeStyles()
    {
        using var w = CreateInMemorySpreadsheet();
        var s1 = w.CreateStyle(
            font: new Font { FontSize = 15, Color = Color.Brown, FontName = FontNameValues.Calibri },
            border: new Border(BorderStyleValues.Double)
        );
        var s2 = w.CreateStyle(
            font: new Font { FontSize = 20, Color = Color.Brown, FontName = FontNameValues.Tahoma },
            numberFormat: new NumberingFormat("0x")
        );

        var s = s1.CreateMergedStyle(s2);

        Assert.True(s.FontId > 0 && s.FontId == s2.FontId);
        Assert.Equal(0, s.FillId);
        Assert.True(s.NumberFormatId > 0 && s.NumberFormatId == s2.NumberFormatId);
        Assert.True(s.BorderId > 0 && s.BorderId == s1.BorderId);
        Assert.True(s.StyleIndex > 0);
    }

    [Fact]
    public void MergeStylesToKnownStyle()
    {
        using var w = CreateInMemorySpreadsheet();
        var sOld = w.CreateStyle(
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

        Assert.Equal(sOld.FontId, s.FontId);
        Assert.Equal(sOld.FillId, s.FillId);
        Assert.Equal(sOld.NumberFormatId, s.NumberFormatId);
        Assert.Equal(sOld.BorderId, s.BorderId);
        Assert.Null(StylesheetProbe.Alignment(s));
        Assert.Null(StylesheetProbe.Alignment(sOld));
        Assert.Equal(sOld.StyleIndex, s.StyleIndex);
    }

    [Fact]
    public void MergeStylesWithNull()
    {
        using var w = CreateInMemorySpreadsheet();
        var sOld = w.CreateStyle(
            font: new Font { FontSize = 20, Color = Color.Brown, FontName = FontNameValues.Tahoma, Bold = true },
            border: new Border(BorderStyleValues.Double),
            numberFormat: new NumberingFormat("0x")
        );

        var s = sOld.CreateMergedStyle(null);

        Assert.Equal(sOld.FontId, s.FontId);
        Assert.Equal(sOld.FillId, s.FillId);
        Assert.Equal(sOld.NumberFormatId, s.NumberFormatId);
        Assert.Equal(sOld.BorderId, s.BorderId);
        Assert.Null(StylesheetProbe.Alignment(s));
        Assert.Null(StylesheetProbe.Alignment(sOld));
        Assert.Equal(sOld.StyleIndex, s.StyleIndex);
    }

    [Fact]
    public void MergeStylesAcrossWorkbooks_UsesTargetStylesheetIds()
    {
        const int firstCustomNumberFormatId = 170;

        using var targetWorkbook = CreateInMemorySpreadsheet();
        using var sourceWorkbook = CreateInMemorySpreadsheet();

        _ = targetWorkbook.CreateStyle(
            font: new Font { FontSize = 11, Color = Color.Blue, FontName = FontNameValues.Calibri },
            fill: new Fill(Color.Yellow),
            numberFormat: new NumberingFormat("#,##0x")
        );

        var targetStyle = targetWorkbook.CreateStyle(border: new Border(BorderStyleValues.Double));
        var sourceStyle = sourceWorkbook.CreateStyle(
            font: new Font { FontSize = 15, Color = Color.Brown, FontName = FontNameValues.Tahoma },
            fill: new Fill(Color.Aqua),
            numberFormat: new NumberingFormat("0x")
        );

        var mergedStyle = targetStyle.CreateMergedStyle(sourceStyle);

        Assert.True(StylesheetProbe.ShareStylesheet(targetStyle, mergedStyle));
        Assert.NotEqual(sourceStyle.FontId, mergedStyle.FontId);
        Assert.NotEqual(sourceStyle.FillId, mergedStyle.FillId);
    Assert.Equal(firstCustomNumberFormatId + 1, mergedStyle.NumberFormatId);

        var sourceFont = StylesheetProbe.Font(sourceStyle);
        var mergedFont = StylesheetProbe.Font(mergedStyle);
        Assert.True(mergedFont.OuterXml.CompareXml(sourceFont.OuterXml));

        var sourceFill = StylesheetProbe.Fill(sourceStyle);
        var mergedFill = StylesheetProbe.Fill(mergedStyle);
        Assert.True(mergedFill.OuterXml.CompareXml(sourceFill.OuterXml));
    }

    [Fact]
    public void MergedStyle_KeepsFontChildrenInSchemaOrder()
    {
        using var workbook = CreateNewSpreadsheet(new MemoryStream());

        // The overlay contributes 'b', which CT_Font requires before the base's 'sz'.
        var sized = workbook.CreateStyle(new Font { FontSize = 33 });
        var bold = workbook.CreateStyle(new Font { Bold = true });

        var merged = sized.CreateMergedStyle(bold);

        var mergedFont = StylesheetProbe.Font(merged);

        Assert.Equal(["b", "sz"], mergedFont.ChildElements.Select(child => child.LocalName));
    }

    [Theory]
    [InlineData("#2A66FF", "FF2A66FF")]
    [InlineData("2A66FF", "FF2A66FF")]
    [InlineData("#FF2A66FF", "FF2A66FF")]
    [InlineData("ff2a66ff", "FF2A66FF")]
    public void ArgbHexColor_IsNormalizedToEightDigitHex(string input, string expected)
    {
        using var workbook = CreateNewSpreadsheet(new MemoryStream());

        var style = workbook.CreateStyle(new Font { ArgbHexColor = input });
        var font = StylesheetProbe.Font(style);

        Assert.Equal(expected, font.Color!.Rgb!.Value);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("#12345")]
    [InlineData("GGGGGG")]
    [InlineData("#2A66FF00FF")]
    public void ArgbHexColor_InvalidValue_Throws(string input)
    {
        Assert.Throws<ArgumentException>(() => new Font { ArgbHexColor = input });
    }

    [Theory]
    [InlineData("#FFFF99", "FFFFFF99")]
    [InlineData("ffff99", "FFFFFF99")]
    public void FillForegroundColor_IsNormalizedToEightDigitHex(string input, string expected)
    {
        using var workbook = CreateNewSpreadsheet(new MemoryStream());

        var style = workbook.CreateStyle(fill: new Fill(input));
        var fill = StylesheetProbe.Fill(style);

        Assert.Equal(expected, fill.PatternFill!.ForegroundColor!.Rgb!.Value);
    }

    [Fact]
    public void CustomNumberFormats_StartFromFirstUserIndexPerWorkbook()
    {
        const int firstCustomNumberFormatId = 170;

        using var workbook1 = CreateInMemorySpreadsheet();
        using var workbook2 = CreateInMemorySpreadsheet();

        var style1 = workbook1.CreateStyle(numberFormat: new NumberingFormat("0x"));
        var style2 = workbook2.CreateStyle(numberFormat: new NumberingFormat("0x"));

        Assert.Equal(firstCustomNumberFormatId, style1.NumberFormatId);
        Assert.Equal(firstCustomNumberFormatId, style2.NumberFormatId);
    }
}