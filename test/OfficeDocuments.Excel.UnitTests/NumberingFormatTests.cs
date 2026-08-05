using OfficeDocuments.Excel.Styles;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.UnitTests;

public class NumberingFormatTests
{
    [Theory]
    [InlineData("General", 0u)]
    [InlineData("0", 1u)]
    [InlineData("0.00", 2u)]
    [InlineData("#,##0", 3u)]
    [InlineData("0%", 9u)]
    [InlineData("@", 49u)]
    public void TryGetBuiltInId_KnownFormatCode_ReturnsExcelBuiltInId(string formatCode, uint expectedId)
    {
        var found = NumberingFormat.TryGetBuiltInId(formatCode, out var numberFormatId);

        Assert.True(found);
        Assert.Equal(expectedId, numberFormatId);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void TryGetBuiltInId_MissingFormatCode_FallsBackToGeneral(string? formatCode)
    {
        var found = NumberingFormat.TryGetBuiltInId(formatCode, out var numberFormatId);

        Assert.True(found);
        Assert.Equal(0u, numberFormatId);
    }

    [Theory]
    [InlineData("0x")]
    [InlineData("#,##0.00#")]
    [InlineData("general")] // built-in lookup is case-sensitive
    public void TryGetBuiltInId_CustomFormatCode_ReturnsFalse(string formatCode)
    {
        var found = NumberingFormat.TryGetBuiltInId(formatCode, out var numberFormatId);

        Assert.False(found);
        Assert.Equal(0u, numberFormatId);
    }

    [Fact]
    public void Constructor_BuiltInFormatCode_AssignsBuiltInId()
    {
        var format = new NumberingFormat("@");

        Assert.Equal(49u, format.Element.NumberFormatId?.Value);
        Assert.Equal("@", format.Element.FormatCode?.Value);
    }

    [Fact]
    public void Constructor_EmptyFormatCode_DefaultsToGeneral()
    {
        var format = new NumberingFormat(string.Empty);

        Assert.Equal("General", format.Element.FormatCode?.Value);
        Assert.Equal(0u, format.Element.NumberFormatId?.Value);
    }

    [Fact]
    public void GetNextCustomId_EmptyCollection_ReturnsFirstUserIndex()
    {
        var numberingFormats = new SpreadsheetLib.NumberingFormats();

        Assert.Equal(170u, NumberingFormat.GetNextCustomId(numberingFormats));
    }

    [Fact]
    public void GetNextCustomId_SkipsPastTheHighestExistingId()
    {
        var numberingFormats = new SpreadsheetLib.NumberingFormats(
            new SpreadsheetLib.NumberingFormat { NumberFormatId = 170u, FormatCode = "0x" },
            new SpreadsheetLib.NumberingFormat { NumberFormatId = 173u, FormatCode = "0y" });

        Assert.Equal(174u, NumberingFormat.GetNextCustomId(numberingFormats));
    }

    [Fact]
    public void GetNextCustomId_IgnoresBuiltInIdsBelowTheUserRange()
    {
        var numberingFormats = new SpreadsheetLib.NumberingFormats(
            new SpreadsheetLib.NumberingFormat { NumberFormatId = 49u, FormatCode = "@" });

        Assert.Equal(170u, NumberingFormat.GetNextCustomId(numberingFormats));
    }

    [Fact]
    public void IsContentSame_ComparesFormatCodeTreatingMissingAsGeneral()
    {
        var format = new NumberingFormat("General");

        Assert.True(format.IsContentSame(new SpreadsheetLib.NumberingFormat()));
        Assert.False(format.IsContentSame(new SpreadsheetLib.NumberingFormat { FormatCode = "0x" }));
    }
}
