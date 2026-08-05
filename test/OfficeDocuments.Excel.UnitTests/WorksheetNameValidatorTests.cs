using OfficeDocuments.Excel.DataClasses;

namespace OfficeDocuments.Excel.UnitTests;

/// <summary>
/// Worksheet-name legality (EXCEL-011 phase 6, blind spot B-2). Before this validation the library
/// wrote whatever it was handed, and the schema constrains the attribute itself — so an over-long
/// or `/`-bearing name produced a workbook Excel offers to repair.
/// </summary>
public class WorksheetNameValidatorTests
{
    private static Exception? Capture(string? name) =>
        Record.Exception(() => WorksheetNameValidator.Validate(name, "name"));

    [Theory]
    [InlineData("Sheet1")]
    [InlineData("Q1 2026")]
    [InlineData("a")]
    [InlineData("Ceník položek")]
    [InlineData("has 'quotes' inside")]
    [InlineData("with & and < and >")]
    public void Accepts_NamesExcelAllows(string name)
    {
        Assert.Null(Capture(name));
    }

    [Fact]
    public void Accepts_ExactlyThirtyOneCharacters()
    {
        Assert.Null(Capture(new string('a', 31)));
        Assert.NotNull(Capture(new string('a', 32)));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Rejects_EmptyNames(string? name)
    {
        Assert.IsType<ArgumentException>(Capture(name));
    }

    /// <summary>
    /// These are exactly the characters that would be ambiguous inside an A1-style reference —
    /// <c>Sheet!A1</c>, <c>[Book]Sheet</c>, <c>A1:B2</c> — which is why Excel forbids them.
    /// </summary>
    [Theory]
    [InlineData(':')]
    [InlineData('\\')]
    [InlineData('/')]
    [InlineData('?')]
    [InlineData('*')]
    [InlineData('[')]
    [InlineData(']')]
    public void Rejects_CharactersExcelForbids(char forbidden)
    {
        var exception = Assert.IsType<ArgumentException>(Capture($"Data{forbidden}2026"));

        Assert.Contains(forbidden.ToString(), exception.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// An apostrophe is legal inside the name but not at either end, where it would collide with
    /// the quoting of a <c>'Sheet Name'!A1</c> reference.
    /// </summary>
    [Theory]
    [InlineData("'Sheet")]
    [InlineData("Sheet'")]
    [InlineData("'")]
    public void Rejects_LeadingOrTrailingApostrophe(string name)
    {
        Assert.IsType<ArgumentException>(Capture(name));
    }

    [Theory]
    [InlineData("History")]
    [InlineData("history")]
    [InlineData("HISTORY")]
    public void Rejects_TheNameExcelReserves(string name)
    {
        Assert.IsType<ArgumentException>(Capture(name));
    }

    [Fact]
    public void Rejects_CharactersXmlCannotRepresent()
    {
        Assert.IsType<ArgumentException>(Capture("Sheet" + (char)0x01));
    }

    /// <summary>The message has to say which name was rejected; the caller may be in a loop.</summary>
    [Fact]
    public void Rejects_WithAMessageNamingTheOffendingValue()
    {
        var tooLong = new string('x', 40);

        var exception = Assert.IsType<ArgumentException>(Capture(tooLong));

        Assert.Contains("31", exception.Message, StringComparison.Ordinal);
        Assert.Contains("40", exception.Message, StringComparison.Ordinal);
    }
}
