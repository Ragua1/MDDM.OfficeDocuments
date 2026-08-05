using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Options;

namespace OfficeDocuments.Excel.UnitTests;

public class DataValidationOptionsTests
{
    [Fact]
    public void List_BuildsQuotedCommaSeparatedFormula()
    {
        var options = DataValidationOptions.List(["A", "B", "C"]);

        Assert.Equal(DataValidationType.List, options.Type);
        Assert.Equal("\"A,B,C\"", options.Formula1);
        Assert.Null(options.Formula2);
        Assert.Null(options.Operator);
    }

    [Fact]
    public void List_EscapesEmbeddedQuotes()
    {
        var options = DataValidationOptions.List(["say \"hi\""]);

        Assert.Equal("\"say \"\"hi\"\"\"", options.Formula1);
    }

    [Fact]
    public void List_SkipsBlankValues()
    {
        var options = DataValidationOptions.List(["A", "", "   ", "B"]);

        Assert.Equal("\"A,B\"", options.Formula1);
    }

    [Fact]
    public void List_NullValues_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => DataValidationOptions.List(null!));
    }

    [Theory]
    [InlineData((object)new string[0])]
    [InlineData((object)new[] { "" })]
    [InlineData((object)new[] { "  ", "\t" })]
    public void List_WithoutUsableValues_Throws(string[] values)
    {
        Assert.Throws<ArgumentException>(() => DataValidationOptions.List(values));
    }

    [Fact]
    public void AllowBlankAndShowDropDown_DefaultToTrue()
    {
        var options = DataValidationOptions.List(["A"]);

        Assert.True(options.AllowBlank);
        Assert.True(options.ShowDropDown);
    }

    [Theory]
    [InlineData(DataValidationType.Whole)]
    [InlineData(DataValidationType.Decimal)]
    [InlineData(DataValidationType.Date)]
    public void Comparison_SetsTypeOperatorAndFormulas(DataValidationType type)
    {
        var options = Create(type, DataValidationOperator.Between, "1", "10");

        Assert.Equal(type, options.Type);
        Assert.Equal(DataValidationOperator.Between, options.Operator);
        Assert.Equal("1", options.Formula1);
        Assert.Equal("10", options.Formula2);
    }

    [Theory]
    [InlineData(DataValidationType.Whole)]
    [InlineData(DataValidationType.Decimal)]
    [InlineData(DataValidationType.Date)]
    public void Comparison_BlankFormula1_Throws(DataValidationType type)
    {
        Assert.Throws<ArgumentException>(() => Create(type, DataValidationOperator.Equal, "   ", null));
    }

    private static DataValidationOptions Create(DataValidationType type, DataValidationOperator @operator, string formula1, string? formula2) =>
        type switch
        {
            DataValidationType.Whole => DataValidationOptions.WholeNumber(@operator, formula1, formula2),
            DataValidationType.Decimal => DataValidationOptions.DecimalNumber(@operator, formula1, formula2),
            DataValidationType.Date => DataValidationOptions.Date(@operator, formula1, formula2),
            _ => throw new ArgumentOutOfRangeException(nameof(type))
        };

    [Theory]
    [InlineData(DataValidationOperator.Between)]
    [InlineData(DataValidationOperator.NotBetween)]
    public void Comparison_BetweenOperatorWithoutFormula2_Throws(DataValidationOperator @operator)
    {
        Assert.Throws<ArgumentException>(() => DataValidationOptions.WholeNumber(@operator, "1"));
    }

    [Theory]
    [InlineData(DataValidationOperator.Equal)]
    [InlineData(DataValidationOperator.GreaterThan)]
    public void Comparison_NonRangeOperatorWithoutFormula2_IsAllowed(DataValidationOperator @operator)
    {
        var options = DataValidationOptions.WholeNumber(@operator, "1");

        Assert.Equal(@operator, options.Operator);
        Assert.Null(options.Formula2);
    }

    [Fact]
    public void Custom_SetsCustomTypeAndFormula()
    {
        var options = DataValidationOptions.Custom("ISNUMBER(A1)");

        Assert.Equal(DataValidationType.Custom, options.Type);
        Assert.Equal("ISNUMBER(A1)", options.Formula1);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void Custom_BlankFormula_Throws(string? formula)
    {
        Assert.Throws<ArgumentException>(() => DataValidationOptions.Custom(formula!));
    }
}
