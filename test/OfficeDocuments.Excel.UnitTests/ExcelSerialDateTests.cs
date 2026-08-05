using OfficeDocuments.Excel.DataClasses;

namespace OfficeDocuments.Excel.UnitTests;

/// <summary>
/// The 1900 date system (EXCEL-011 phase 6, blind spot B-6).
/// <para>
/// This is the bug class a round-trip test can never find. The library used
/// <see cref="DateTime.ToOADate"/> to write and <c>FromOADate</c> to read, which are exact
/// inverses — so its own files always read back correctly, while Excel read every date before
/// March 1900 as one day later. The assertions here are therefore stated against **Excel's**
/// serials, taken from the format, not against whatever the code happens to produce.
/// </para>
/// </summary>
public class ExcelSerialDateTests
{
    /// <summary>
    /// The reference points. Serial 60 is missing from this table on purpose: Excel assigns it to
    /// 29 February 1900, a day that did not exist in a year that was not a leap year. The bug came
    /// from Lotus 1-2-3 and is preserved deliberately for file compatibility.
    /// </summary>
    [Theory]
    [InlineData(1900, 1, 1, 1)]      // Excel's epoch
    [InlineData(1900, 1, 2, 2)]
    [InlineData(1900, 2, 28, 59)]    // the last day before the phantom
    [InlineData(1900, 3, 1, 61)]     // 60 is skipped; from here on OLE Automation agrees
    [InlineData(1900, 3, 2, 62)]
    [InlineData(2026, 7, 28, 46231)]
    [InlineData(9999, 12, 31, 2958465)]
    public void ToSerial_MatchesExcel(int year, int month, int day, double expected)
    {
        Assert.Equal(expected, ExcelSerialDate.ToSerial(new DateTime(year, month, day)));
    }

    /// <summary>
    /// The specific regression: below the phantom day, the OLE Automation serial is one too high.
    /// Above it, the two systems agree and this library's old behaviour was already correct.
    /// </summary>
    [Fact]
    public void ToSerial_DivergesFromOaDateOnlyBeforeMarch1900()
    {
        var beforeThePhantomDay = new DateTime(1900, 2, 28);
        var afterThePhantomDay = new DateTime(1900, 3, 1);

        Assert.Equal(beforeThePhantomDay.ToOADate() - 1, ExcelSerialDate.ToSerial(beforeThePhantomDay));
        Assert.Equal(afterThePhantomDay.ToOADate(), ExcelSerialDate.ToSerial(afterThePhantomDay));
    }

    [Fact]
    public void ToSerial_KeepsTheTimeOfDayAsTheFraction()
    {
        Assert.Equal(46231.5, ExcelSerialDate.ToSerial(new DateTime(2026, 7, 28, 12, 0, 0)));
        Assert.Equal(1.25, ExcelSerialDate.ToSerial(new DateTime(1900, 1, 1, 6, 0, 0)));
    }

    /// <summary>
    /// Excel's numbering starts at 1 January 1900 and it has no way to show anything earlier as a
    /// date. Writing a smaller serial produces a cell Excel renders as an error, so the library
    /// refuses rather than silently emitting one.
    /// </summary>
    [Theory]
    [InlineData(1899, 12, 31)]
    [InlineData(1800, 1, 1)]
    public void ToSerial_RejectsDatesExcelCannotRepresent(int year, int month, int day)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => ExcelSerialDate.ToSerial(new DateTime(year, month, day)));
    }

    /// <summary>
    /// <c>DateTime.MinValue.ToOADate()</c> quietly returns 0 rather than throwing, so the old code
    /// turned "no date" into 30 December 1899 without a word. It has to be refused explicitly.
    /// </summary>
    [Fact]
    public void ToSerial_RejectsDateTimeMinValue()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => ExcelSerialDate.ToSerial(DateTime.MinValue));
    }

    [Theory]
    [InlineData(1, 1900, 1, 1)]
    [InlineData(59, 1900, 2, 28)]
    [InlineData(61, 1900, 3, 1)]
    [InlineData(46231, 2026, 7, 28)]
    public void FromSerial_MatchesExcel(double serial, int year, int month, int day)
    {
        Assert.Equal(new DateTime(year, month, day), ExcelSerialDate.FromSerial(serial));
    }

    /// <summary>
    /// Serial 60 arrives only from a foreign producer, and it denotes a date that does not exist.
    /// Reading is permissive — the alternative is refusing to open the file — and resolves it to
    /// the next real day.
    /// </summary>
    [Fact]
    public void FromSerial_ResolvesThePhantomLeapDayToTheFollowingDay()
    {
        Assert.Equal(new DateTime(1900, 3, 1), ExcelSerialDate.FromSerial(60));
    }

    [Theory]
    [InlineData(1900, 1, 1)]
    [InlineData(1900, 2, 28)]
    [InlineData(1900, 3, 1)]
    [InlineData(1999, 12, 31)]
    [InlineData(2026, 7, 28)]
    public void ToSerial_AndBack_IsLossless(int year, int month, int day)
    {
        var original = new DateTime(year, month, day);

        Assert.Equal(original, ExcelSerialDate.FromSerial(ExcelSerialDate.ToSerial(original)));
    }
}
