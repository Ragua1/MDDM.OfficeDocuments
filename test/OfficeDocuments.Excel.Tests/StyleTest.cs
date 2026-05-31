using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Styles;
using OfficeDocuments.Excel.Extensions;
using Color = System.Drawing.Color;

namespace OfficeDocuments.Excel.Tests;

public class StyleTest : SpreadsheetTestBase
{
    [Fact]
    public void BasicStyle()
    {
        var filePath = GetFilepath("doc1.xlsx");
        using var w = CreateNewSpreadsheet(filePath);
        var s = w.CreateStyle();

        Assert.NotNull(s.Element);
        Assert.Equal(0, s.FontId);
        Assert.Equal(0, s.FillId);
        Assert.Equal(0, s.NumberFormatId);
        Assert.Equal(0, s.BorderId);
        Assert.Equal(0U, s.StyleIndex);
    }

    [Fact]
    public void SpecificFontStyle()
    {
        var filePath = GetFilepath("doc2.xlsx");
        using var w = CreateNewSpreadsheet(filePath);
        var s = w.CreateStyle(
            new Font { FontSize = 15, Color = Color.Blue, FontName = FontNameValues.Tahoma, Bold = true, Italic = true, Underline = UnderlineValues.Double }
        );

        Assert.NotNull(s.Element);
        Assert.True(s.FontId > 0);
        Assert.Equal(0, s.FillId);
        Assert.Equal(0, s.NumberFormatId);
        Assert.Equal(0, s.BorderId);
        Assert.True(s.StyleIndex > 0);
    }

    [Fact]
    public void SpecificFillStyle()
    {
        var filePath = GetFilepath("doc3.xlsx");
        using var w = CreateNewSpreadsheet(filePath);
        var s = w.CreateStyle(
            fill: new Fill(Color.Blue, Color.White)
        );

        Assert.NotNull(s.Element);
        Assert.Equal(0, s.FontId);
        Assert.True(s.FillId > 0);
        Assert.Equal(0, s.NumberFormatId);
        Assert.Equal(0, s.BorderId);
        Assert.True(s.StyleIndex > 0);
    }

    [Fact]
    public void SpecificBorderStyle()
    {
        var filePath = GetFilepath("doc4.xlsx");
        using var w = CreateNewSpreadsheet(filePath);
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

        Assert.NotNull(s.Element);
        Assert.Equal(0, s.FontId);
        Assert.Equal(0, s.FillId);
        Assert.Equal(0, s.NumberFormatId);
        Assert.True(s.BorderId > 0);
        Assert.True(s.StyleIndex > 0);
    }

    [Fact]
    public void SpecificBorderStyle1()
    {
        var filePath = GetFilepath("doc5.xlsx");
        using var w = CreateNewSpreadsheet(filePath);
        var s = w.CreateStyle(
            border: new Border(BorderStyleValues.Medium)
        );

        Assert.NotNull(s.Element);
        Assert.Equal(0, s.FontId);
        Assert.Equal(0, s.FillId);
        Assert.Equal(0, s.NumberFormatId);
        Assert.True(s.BorderId > 0);
        Assert.True(s.StyleIndex > 0);
    }

    [Fact]
    public void SpecificNumberFormatStyle()
    {
        var filePath = GetFilepath("doc6.xlsx");
        using var w = CreateNewSpreadsheet(filePath);
        var s = w.CreateStyle(
            numberFormat: new NumberingFormat("@")
        );

        Assert.NotNull(s.Element);
        Assert.Equal(0, s.FontId);
        Assert.Equal(0, s.FillId);
        Assert.Equal(49, s.NumberFormatId);
        Assert.Equal(0, s.BorderId);
        Assert.True(s.StyleIndex > 0);
    }

    [Fact]
    public void SpecificAlignmentStyle()
    {
        var filePath = GetFilepath("doc7.xlsx");
        using var w = CreateNewSpreadsheet(filePath);
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

        Assert.NotNull(s.Element);
        Assert.Equal(0, s.FontId);
        Assert.Equal(0, s.FillId);
        Assert.Equal(0, s.NumberFormatId);
        Assert.Equal(0, s.BorderId);
        Assert.NotNull(s.Element.Alignment);
        Assert.True(s.StyleIndex > 0);
    }

    [Fact]
    public void MergeStyles()
    {
        var filePath = GetFilepath("doc8.xlsx");
        using var w = CreateNewSpreadsheet(filePath);
        var s1 = w.CreateStyle(
            font: new Font { FontSize = 15, Color = Color.Brown, FontName = FontNameValues.Calibri },
            border: new Border(BorderStyleValues.Double)
        );
        var s2 = w.CreateStyle(
            font: new Font { FontSize = 20, Color = Color.Brown, FontName = FontNameValues.Tahoma },
            numberFormat: new NumberingFormat("0x")
        );

        var s = s1.CreateMergedStyle(s2);

        Assert.NotNull(s.Element);
        Assert.True(s.FontId > 0 && s.FontId == s2.FontId);
        Assert.Equal(0, s.FillId);
        Assert.True(s.NumberFormatId > 0 && s.NumberFormatId == s2.NumberFormatId);
        Assert.True(s.BorderId > 0 && s.BorderId == s1.BorderId);
        Assert.True(s.StyleIndex > 0);
    }

    [Fact]
    public void MergeStylesToKnownStyle()
    {
        var filePath = GetFilepath("doc9.xlsx");
        using var w = CreateNewSpreadsheet(filePath);
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

        Assert.NotNull(s.Element);
        Assert.Equal(sOld.FontId, s.FontId);
        Assert.Equal(sOld.FillId, s.FillId);
        Assert.Equal(sOld.NumberFormatId, s.NumberFormatId);
        Assert.Equal(sOld.BorderId, s.BorderId);
        Assert.Null(s.Element.Alignment);
        Assert.Null(sOld.Element.Alignment);
        Assert.Equal(sOld.StyleIndex, s.StyleIndex);
    }

    [Fact]
    public void MergeStylesWithNull()
    {
        var filePath = GetFilepath("doc10.xlsx");
        using var w = CreateNewSpreadsheet(filePath);
        var sOld = w.CreateStyle(
            font: new Font { FontSize = 20, Color = Color.Brown, FontName = FontNameValues.Tahoma, Bold = true },
            border: new Border(BorderStyleValues.Double),
            numberFormat: new NumberingFormat("0x")
        );

        var s = sOld.CreateMergedStyle(null);

        Assert.NotNull(s.Element);
        Assert.Equal(sOld.FontId, s.FontId);
        Assert.Equal(sOld.FillId, s.FillId);
        Assert.Equal(sOld.NumberFormatId, s.NumberFormatId);
        Assert.Equal(sOld.BorderId, s.BorderId);
        Assert.Null(s.Element.Alignment);
        Assert.Null(sOld.Element.Alignment);
        Assert.Equal(sOld.StyleIndex, s.StyleIndex);
    }

    [Fact]
    public void MergeStylesAcrossWorkbooks_UsesTargetStylesheetIds()
    {
        const int firstCustomNumberFormatId = 170;

        using var targetWorkbook = CreateNewSpreadsheet(GetFilepath("doc11-target.xlsx"));
        using var sourceWorkbook = CreateNewSpreadsheet(GetFilepath("doc11-source.xlsx"));

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

        Assert.Same(targetStyle.Stylesheet, mergedStyle.Stylesheet);
        Assert.NotEqual(sourceStyle.FontId, mergedStyle.FontId);
        Assert.NotEqual(sourceStyle.FillId, mergedStyle.FillId);
    Assert.Equal(firstCustomNumberFormatId + 1, mergedStyle.NumberFormatId);

        var sourceFont = sourceStyle.Stylesheet.Fonts!.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ElementAt(sourceStyle.FontId);
        var mergedFont = mergedStyle.Stylesheet.Fonts!.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ElementAt(mergedStyle.FontId);
        Assert.True(mergedFont.OuterXml.CompareXml(sourceFont.OuterXml));

        var sourceFill = sourceStyle.Stylesheet.Fills!.Elements<DocumentFormat.OpenXml.Spreadsheet.Fill>().ElementAt(sourceStyle.FillId);
        var mergedFill = mergedStyle.Stylesheet.Fills!.Elements<DocumentFormat.OpenXml.Spreadsheet.Fill>().ElementAt(mergedStyle.FillId);
        Assert.True(mergedFill.OuterXml.CompareXml(sourceFill.OuterXml));
    }

    [Fact]
    public void CustomNumberFormats_StartFromFirstUserIndexPerWorkbook()
    {
        const int firstCustomNumberFormatId = 170;

        using var workbook1 = CreateNewSpreadsheet(GetFilepath("doc12.xlsx"));
        using var workbook2 = CreateNewSpreadsheet(GetFilepath("doc13.xlsx"));

        var style1 = workbook1.CreateStyle(numberFormat: new NumberingFormat("0x"));
        var style2 = workbook2.CreateStyle(numberFormat: new NumberingFormat("0x"));

        Assert.Equal(firstCustomNumberFormatId, style1.NumberFormatId);
        Assert.Equal(firstCustomNumberFormatId, style2.NumberFormatId);
    }
}