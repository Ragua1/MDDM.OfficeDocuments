using System.Globalization;
using OfficeDocuments.Excel.TestKit;

namespace OfficeDocuments.Excel.IntegrationTests;

/// <summary>
/// Dates through the public API (EXCEL-011 phase 6, blind spot B-6).
/// <para>
/// These assert on the **serial actually written**, not only on what reads back. That distinction
/// is the entire lesson of this blind spot: the library used to write with
/// <see cref="DateTime.ToOADate"/> and read with <c>FromOADate</c>, so its own round trip was
/// perfect while Excel read every date before March 1900 one day late. A test that only round
/// trips cannot see a bug where both halves are wrong in the same direction.
/// </para>
/// </summary>
public class DateValueTests : SpreadsheetTestBase
{
    private static string WrittenSerial(DateTime value)
    {
        using var workbook = CreateInMemorySpreadsheet();
        var cell = workbook.AddWorksheet("Sheet 1").AddCell(1u, 1u, value);
        return cell.GetStringValue() ?? string.Empty;
    }

    [Theory]
    [InlineData(1900, 1, 1, "1")]
    [InlineData(1900, 2, 28, "59")]
    [InlineData(1900, 3, 1, "61")]
    [InlineData(2026, 7, 28, "46231")]
    public void SetValue_WritesTheSerialExcelExpects(int year, int month, int day, string expected)
    {
        Assert.Equal(expected, WrittenSerial(new DateTime(year, month, day)));
    }

    /// <summary>
    /// Serial 60 is Excel's 29 February 1900, a day that never existed. Nothing this library
    /// writes may land on it, because there is no real date it could mean.
    /// </summary>
    [Fact]
    public void SetValue_NeverWritesThePhantomLeapDaySerial()
    {
        Assert.NotEqual("60", WrittenSerial(new DateTime(1900, 2, 28)));
        Assert.NotEqual("60", WrittenSerial(new DateTime(1900, 3, 1)));
    }

    [Fact]
    public void SetValue_KeepsTheTimeOfDay()
    {
        var serial = double.Parse(WrittenSerial(new DateTime(2026, 7, 28, 18, 0, 0)), CultureInfo.InvariantCulture);

        Assert.Equal(46231.75, serial);
    }

    /// <summary>
    /// Excel's serial numbering starts at 1 January 1900. Anything earlier used to be written as a
    /// zero or negative serial, which Excel renders as an error rather than a date — and
    /// <c>DateTime.MinValue</c> was the worst case, because <c>ToOADate</c> quietly maps it to 0
    /// instead of throwing, so "no date" silently became 30 December 1899.
    /// </summary>
    [Theory]
    [InlineData(1899, 12, 31)]
    [InlineData(1850, 6, 15)]
    public void SetValue_RejectsDatesBeforeExcelsEpoch(int year, int month, int day)
    {
        using var workbook = CreateInMemorySpreadsheet();
        var worksheet = workbook.AddWorksheet("Sheet 1");

        Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.AddCell(1u, 1u, new DateTime(year, month, day)));
    }

    [Fact]
    public void SetValue_RejectsDateTimeMinValue()
    {
        using var workbook = CreateInMemorySpreadsheet();
        var worksheet = workbook.AddWorksheet("Sheet 1");

        Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.AddCell(1u, 1u, DateTime.MinValue));
    }

    [Theory]
    [InlineData(1900, 1, 1)]
    [InlineData(1900, 2, 28)]
    [InlineData(1900, 3, 1)]
    [InlineData(2026, 7, 28)]
    [InlineData(9999, 12, 31)]
    public void SetValue_AndReadBack_PreservesTheDate(int year, int month, int day)
    {
        var written = new DateTime(year, month, day);

        using var workbook = CreateInMemorySpreadsheet(out var stream);
        workbook.AddWorksheet("Sheet 1").AddCell(1u, 1u, written);
        workbook.Close();

        stream.Position = 0;
        using var reopened = OpenExistingSpreadsheet(stream);
        Assert.True(reopened.GetWorksheet("Sheet 1")!.GetCell(1u, 1u)!.TryGetValue(out DateTime readBack));
        Assert.Equal(written, readBack);
    }

    /// <summary>
    /// Reading is permissive where writing is strict, because the read path also sees files this
    /// library did not write. A foreign producer's serial 60 resolves to the following real day.
    /// </summary>
    [Fact]
    public void TryGetValue_ResolvesAForeignPhantomLeapDaySerial()
    {
        using var workbook = CreateInMemorySpreadsheet(out var stream);
        // Written as a number, so the value reaches the file as the raw serial 60.
        workbook.AddWorksheet("Sheet 1").AddCell(1u, 1u, 60);
        workbook.Close();

        stream.Position = 0;
        using var reopened = OpenExistingSpreadsheet(stream);
        Assert.True(reopened.GetWorksheet("Sheet 1")!.GetCell(1u, 1u)!.TryGetValue(out DateTime readBack));
        Assert.Equal(new DateTime(1900, 3, 1), readBack);
    }
}
