using Color = System.Drawing.Color;

namespace OfficeDocuments.Excel.UnitTests;

/// <summary>
/// Pure-function cover for the colour handling. The rendered-stylesheet side of this lives in the
/// integration tier; here the concern is only the string contract, including the invalid inputs
/// that used to be written straight into the <c>rgb</c> hexBinary attribute (EXCEL-011 phase 1).
/// </summary>
public class UtilsColorTests
{
    [Fact]
    public void ArgbHexConverter_ProducesEightUppercaseDigitsInArgbOrder()
    {
        Assert.Equal("FFFF0000", Utils.ArgbHexConverter(Color.Red));
        Assert.Equal("FF0000FF", Utils.ArgbHexConverter(Color.Blue));
        Assert.Equal("00000000", Utils.ArgbHexConverter(Color.FromArgb(0, 0, 0, 0)));

        // Color.Transparent is ARGB(0, 255, 255, 255), not a zeroed colour — a classic trap
        // when a caller assumes "transparent" means "no colour bits set".
        Assert.Equal("00FFFFFF", Utils.ArgbHexConverter(Color.Transparent));
    }

    [Fact]
    public void ArgbHexConverter_PreservesAlpha()
    {
        Assert.Equal("802A66FF", Utils.ArgbHexConverter(Color.FromArgb(0x80, 0x2A, 0x66, 0xFF)));
    }

    [Theory]
    [InlineData("#2A66FF", "FF2A66FF")]
    [InlineData("2A66FF", "FF2A66FF")]
    [InlineData("#FF2A66FF", "FF2A66FF")]
    [InlineData("FF2A66FF", "FF2A66FF")]
    [InlineData("ff2a66ff", "FF2A66FF")]
    [InlineData("#ffff99", "FFFFFF99")]
    [InlineData("  #2A66FF  ", "FF2A66FF")]
    [InlineData("00000000", "00000000")]
    public void NormalizeArgbHex_AcceptsRgbAndArgbWithOptionalHash(string input, string expected)
    {
        Assert.Equal(expected, Utils.NormalizeArgbHex(input, "color"));
    }

    [Fact]
    public void NormalizeArgbHex_IsIdempotent()
    {
        var once = Utils.NormalizeArgbHex("#2a66ff", "color");

        Assert.Equal(once, Utils.NormalizeArgbHex(once, "color"));
    }

    [Fact]
    public void NormalizeArgbHex_RoundTripsArgbHexConverter()
    {
        var converted = Utils.ArgbHexConverter(Color.RebeccaPurple);

        Assert.Equal(converted, Utils.NormalizeArgbHex(converted, "color"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("#")]
    [InlineData("12345")]      // too short
    [InlineData("1234567")]    // between the two valid lengths
    [InlineData("123456789")]  // too long
    [InlineData("GGGGGG")]     // not hexadecimal
    [InlineData("#2A66F!")]
    [InlineData("0x2A66FF")]
    public void NormalizeArgbHex_InvalidValue_ThrowsArgumentExceptionNamingTheParameter(string input)
    {
        var exception = Assert.Throws<ArgumentException>(() => Utils.NormalizeArgbHex(input, "foregroundColor"));

        Assert.Equal("foregroundColor", exception.ParamName);
    }

    [Fact]
    public void NormalizeArgbHex_Null_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => Utils.NormalizeArgbHex(null!, "color"));
    }
}
