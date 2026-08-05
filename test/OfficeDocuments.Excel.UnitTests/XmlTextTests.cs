using OfficeDocuments.Excel.DataClasses;

namespace OfficeDocuments.Excel.UnitTests;

/// <summary>
/// Characters XML cannot carry (EXCEL-011 phase 6, blind spot B-4).
/// <para>
/// The distinction that matters here is between "needs escaping" and "cannot be written". `&amp;`
/// and `&lt;` are ordinary text that the SDK escapes on the way out; the C0 control characters
/// have no XML spelling at all, in any encoding, escaped or not. Only the second group belongs in
/// this guard — rejecting the first would break every workbook containing an ampersand.
/// </para>
/// </summary>
public class XmlTextTests
{
    [Theory]
    [InlineData("plain text")]
    [InlineData("markup <b>&amp;</b> \"quoted\" 'apostrophes'")]
    [InlineData("tab\tnewline\nreturn\r")]     // the three control characters XML does allow
    [InlineData("accented ěščřžýáíé")]
    [InlineData("emoji 😀 as a surrogate pair")]
    public void IndexOfIllegalCharacter_AcceptsRepresentableText(string value)
    {
        Assert.Equal(-1, XmlText.IndexOfIllegalCharacter(value));
    }

    [Theory]
    [InlineData(0x00)]
    [InlineData(0x01)]
    [InlineData(0x08)]
    [InlineData(0x0B)]
    [InlineData(0x0C)]
    [InlineData(0x0E)]
    [InlineData(0x1F)]
    public void IndexOfIllegalCharacter_RejectsControlCharacters(int codePoint)
    {
        Assert.Equal(2, XmlText.IndexOfIllegalCharacter("ab" + (char)codePoint + "cd"));
    }

    /// <summary>
    /// A lone surrogate is not a character, only half of one. It survives inside a .NET string but
    /// cannot be encoded, so it has to be caught here rather than at serialization time.
    /// </summary>
    [Fact]
    public void IndexOfIllegalCharacter_RejectsAnUnpairedSurrogate()
    {
        Assert.Equal(1, XmlText.IndexOfIllegalCharacter("a\uD83Db"));
    }

    [Fact]
    public void EnsureRepresentable_SaysWhichCharacterAndWhere()
    {
        var exception = Assert.Throws<ArgumentException>(
            () => XmlText.EnsureRepresentable("ab" + (char)0x01, "value", "The value for cell 'A1'"));

        Assert.Contains("A1", exception.Message, StringComparison.Ordinal);
        Assert.Contains("U+0001", exception.Message, StringComparison.Ordinal);
        Assert.Contains("index 2", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void EnsureRepresentable_PassesTextThatOnlyNeedsEscaping()
    {
        XmlText.EnsureRepresentable("<a>&b\"c'd", "value", "The value for cell 'A1'");
    }
}
