using DocumentFormat.OpenXml.Spreadsheet;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Interfaces;
using OfficeDocuments.Excel.Options;
using Color = System.Drawing.Color;

namespace OfficeDocuments.Excel.UnitTests;

public class ConditionalFormattingOptionsTests
{
    private static readonly IStyle Style = new StubStyle();

    [Theory]
    [InlineData(ConditionalFormattingType.GreaterThan)]
    [InlineData(ConditionalFormattingType.LessThan)]
    public void Threshold_SetsTypeFormulaAndStyle(ConditionalFormattingType type)
    {
        var options = CreateThreshold(type, "10", Style);

        Assert.Equal(type, options.Type);
        Assert.Equal("10", options.Formula);
        Assert.Same(Style, options.Style);
        Assert.Null(options.Text);
    }

    [Theory]
    [InlineData(ConditionalFormattingType.GreaterThan)]
    [InlineData(ConditionalFormattingType.LessThan)]
    public void Threshold_BlankFormula_Throws(ConditionalFormattingType type)
    {
        Assert.Throws<ArgumentException>(() => CreateThreshold(type, "  ", Style));
    }

    [Theory]
    [InlineData(ConditionalFormattingType.GreaterThan)]
    [InlineData(ConditionalFormattingType.LessThan)]
    public void Threshold_NullStyle_Throws(ConditionalFormattingType type)
    {
        Assert.Throws<ArgumentNullException>(() => CreateThreshold(type, "10", null!));
    }

    [Fact]
    public void ContainsText_SetsTypeTextAndStyle()
    {
        var options = ConditionalFormattingOptions.ContainsText("warn", Style);

        Assert.Equal(ConditionalFormattingType.ContainsText, options.Type);
        Assert.Equal("warn", options.Text);
        Assert.Same(Style, options.Style);
        Assert.Null(options.Formula);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void ContainsText_BlankText_Throws(string? text)
    {
        Assert.Throws<ArgumentException>(() => ConditionalFormattingOptions.ContainsText(text!, Style));
    }

    [Fact]
    public void ContainsText_NullStyle_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => ConditionalFormattingOptions.ContainsText("warn", null!));
    }

    [Fact]
    public void DuplicateValues_SetsTypeAndStyle()
    {
        var options = ConditionalFormattingOptions.DuplicateValues(Style);

        Assert.Equal(ConditionalFormattingType.DuplicateValues, options.Type);
        Assert.Same(Style, options.Style);
    }

    [Fact]
    public void DuplicateValues_NullStyle_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => ConditionalFormattingOptions.DuplicateValues(null!));
    }

    [Fact]
    public void TwoColorScale_CarriesBothColorsAndNeedsNoStyle()
    {
        var options = ConditionalFormattingOptions.TwoColorScale(Color.Red, Color.Green);

        Assert.Equal(ConditionalFormattingType.TwoColorScale, options.Type);
        Assert.Equal(Color.Red, options.MinimumColor);
        Assert.Equal(Color.Green, options.MaximumColor);
        Assert.Null(options.Style);
    }

    private static ConditionalFormattingOptions CreateThreshold(ConditionalFormattingType type, string formula, IStyle style) =>
        type switch
        {
            ConditionalFormattingType.GreaterThan => ConditionalFormattingOptions.GreaterThan(formula, style),
            ConditionalFormattingType.LessThan => ConditionalFormattingOptions.LessThan(formula, style),
            _ => throw new ArgumentOutOfRangeException(nameof(type))
        };

    /// <summary>
    /// The options factories only store the style reference, so a stub is enough to exercise
    /// their guard clauses without a workbook.
    /// </summary>
    private sealed class StubStyle : IStyle
    {
#pragma warning disable CS0618 // exercising the transitional raw-OpenXml surface is unavoidable here
        public Stylesheet Stylesheet => throw new NotSupportedException();
        public CellFormat Element => throw new NotSupportedException();
#pragma warning restore CS0618
        public uint StyleIndex => 0;
        public int FontId => 0;
        public int FillId => 0;
        public int BorderId => 0;
        public int NumberFormatId => 0;
        public IStyle CreateMergedStyle(IStyle? style) => this;
    }
}
