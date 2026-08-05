using OfficeDocuments.Excel.TestKit;

namespace OfficeDocuments.Excel.IntegrationTests;

/// <summary>
/// Values a cell cannot legally hold, and values it can hold but that look as if it could not
/// (EXCEL-011 phase 6, blind spots B-3, B-4 and B-5).
/// <para>
/// Two families, and telling them apart is the whole point. Markup characters are ordinary text
/// that the SDK escapes on the way out and unescapes on the way back — those must keep working
/// untouched. Non-finite numbers and C0 control characters have no representation in the format at
/// all, and used to reach the file: the first silently, the second as an exception thrown from
/// <c>Close()</c> long after the offending call.
/// </para>
/// </summary>
public class CellValueLimitsTests : SpreadsheetTestBase
{
    /// <summary>
    /// The one no gate caught. <c>&lt;v&gt;NaN&lt;/v&gt;</c> in a cell marked <c>t="n"</c> passes
    /// the schema validator, because <c>v</c> is declared as a string and "this must be a number"
    /// comes from the cell's type attribute — a semantic rule the validator does not evaluate. It
    /// also survives a round trip, because <c>double.Parse</c> reads "NaN" back happily. Only a
    /// check at the point of assignment stops it.
    /// </summary>
    [Fact]
    public void SetValue_RejectsNonFiniteDoubles()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var worksheet = workbook.AddWorksheet("Sheet 1");

        Assert.Throws<ArgumentException>(() => worksheet.AddCell(1u, 1u, double.NaN));
        Assert.Throws<ArgumentException>(() => worksheet.AddCell(2u, 1u, double.PositiveInfinity));
        Assert.Throws<ArgumentException>(() => worksheet.AddCell(3u, 1u, double.NegativeInfinity));
    }

    [Fact]
    public void SetValue_RejectsNonFiniteFloats()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var worksheet = workbook.AddWorksheet("Sheet 1");

        Assert.Throws<ArgumentException>(() => worksheet.AddCell(1u, 1u, float.NaN));
        Assert.Throws<ArgumentException>(() => worksheet.AddCell(2u, 1u, float.PositiveInfinity));
    }

    /// <summary>The message has to name the cell — a bulk import hits this on one row out of many.</summary>
    [Fact]
    public void SetValue_NonFiniteMessageNamesTheCell()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var worksheet = workbook.AddWorksheet("Sheet 1");

        var exception = Assert.Throws<ArgumentException>(() => worksheet.AddCell(3u, 7u, double.NaN));

        Assert.Contains("C7", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void SetValue_AcceptsFiniteExtremes()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var worksheet = workbook.AddWorksheet("Sheet 1");

        worksheet.AddCell(1u, 1u, 0d);
        worksheet.AddCell(2u, 1u, -0d);
        worksheet.AddCell(3u, 1u, double.Epsilon);
        worksheet.AddCell(4u, 1u, decimal.MaxValue);

        Assert.True(worksheet.GetCell(3u, 1u)!.TryGetValue(out double epsilon));
        Assert.Equal(double.Epsilon, epsilon);
    }

    /// <summary>
    /// Before this, the value was accepted here and the SDK threw during <c>Close()</c> — the whole
    /// document lost, with a message that named neither the sheet nor the cell. The fix is not that
    /// it throws, but *where*.
    /// </summary>
    [Theory]
    [InlineData(0x00)]
    [InlineData(0x01)]
    [InlineData(0x0B)]
    [InlineData(0x1F)]
    public void SetValue_RejectsControlCharactersAtTheCallThatCausesThem(int codePoint)
    {
        using var workbook = CreateInMemorySpreadsheet();
        var worksheet = workbook.AddWorksheet("Sheet 1");

        var exception = Assert.Throws<ArgumentException>(
            () => worksheet.AddCell(2u, 3u, "before" + (char)codePoint + "after"));

        Assert.Contains("B3", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void SetValue_AcceptsNewlinesAndTabs()
    {
        const string value = "line\nnext\tcolumn";

        using var workbook = CreateInMemorySpreadsheet(out var stream);
        workbook.AddWorksheet("Sheet 1").AddCell(1u, 1u, value);
        workbook.Close();

        stream.Position = 0;
        using var reopened = OpenExistingSpreadsheet(stream);
        Assert.Equal(value, reopened.GetWorksheet("Sheet 1")!.GetCell(1u, 1u)!.GetStringValue());
    }

    /// <summary>
    /// A carriage return does not survive, and that is XML rather than this library: a conforming
    /// parser is required to normalize <c>\r\n</c> and a lone <c>\r</c> to <c>\n</c> before the
    /// application ever sees the text. Surviving would require writing it as the character
    /// reference <c>&amp;#xD;</c>, which the SDK does not do.
    /// <para>
    /// Pinned rather than fixed, because Excel agrees: an in-cell line break is <c>\n</c>, which is
    /// what Alt+Enter inserts. Callers handing the library Windows line endings should know the
    /// value they read back will differ from the one they wrote.
    /// </para>
    /// </summary>
    [Fact]
    public void SetValue_NormalizesWindowsLineEndingsToNewlines()
    {
        using var workbook = CreateInMemorySpreadsheet(out var stream);
        workbook.AddWorksheet("Sheet 1").AddCell(1u, 1u, "line\r\nnext");
        workbook.Close();

        stream.Position = 0;
        using var reopened = OpenExistingSpreadsheet(stream);
        Assert.Equal("line\nnext", reopened.GetWorksheet("Sheet 1")!.GetCell(1u, 1u)!.GetStringValue());
    }

    /// <summary>
    /// The escaping side, which already worked and now cannot regress silently. A value containing
    /// every character XML gives special meaning has to come back byte for byte.
    /// </summary>
    [Fact]
    public void SetValue_RoundTripsMarkupCharactersUnchanged()
    {
        const string value = "<tag attr=\"v\"> & 'quoted' </tag>";

        using var workbook = CreateInMemorySpreadsheet(out var stream);
        workbook.AddWorksheet("Sheet 1").AddCell(1u, 1u, value);
        workbook.Close();

        stream.Position = 0;
        using var reopened = OpenExistingSpreadsheet(stream);
        Assert.Equal(value, reopened.GetWorksheet("Sheet 1")!.GetCell(1u, 1u)!.GetStringValue());
    }

    [Fact]
    public void SetComment_RejectsControlCharactersInTextAndAuthor()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var worksheet = workbook.AddWorksheet("Sheet 1");
        var cell = worksheet.AddCell(1u, 1u, "value");

        Assert.Throws<ArgumentException>(() => cell.SetComment("bad" + (char)0x01, "author"));
        Assert.Throws<ArgumentException>(() => cell.SetComment("fine", "bad" + (char)0x01));
    }

    [Fact]
    public void SetComment_RoundTripsMarkupCharactersUnchanged()
    {
        const string comment = "compare a < b && c > d";

        using var workbook = CreateInMemorySpreadsheet(out var stream);
        workbook.AddWorksheet("Sheet 1").AddCell(1u, 1u, "value").SetComment(comment, "R&D");
        workbook.Close();

        stream.Position = 0;
        using var reopened = OpenExistingSpreadsheet(stream);
        Assert.Equal(comment, reopened.GetWorksheet("Sheet 1")!.GetCell(1u, 1u)!.GetComment());
    }
}
