namespace OfficeDocuments.Excel.DataClasses;

/// <summary>
/// Converts between <see cref="DateTime"/> and the serial numbers SpreadsheetML stores dates as.
/// <para>
/// This is deliberately <b>not</b> <see cref="DateTime.ToOADate"/>, which is the obvious choice and
/// the wrong one. The two systems agree from 1 March 1900 onward and differ by exactly one day
/// before it, for two compounding reasons:
/// </para>
/// <list type="bullet">
///   <item>
///     OLE Automation counts from 30 December 1899, so its serial 1 is 31 December 1899. Excel
///     counts from 1 January 1900, which is its serial 1.
///   </item>
///   <item>
///     Excel's 1900 date system contains a day that never existed — serial 60 is 29 February 1900,
///     a leap day in a year that had none. It was inherited from Lotus 1-2-3 deliberately, for
///     file compatibility, and it can never be removed. The phantom day re-aligns the two systems
///     from 1 March 1900 onward, which is exactly why the divergence is invisible for essentially
///     every date anyone writes.
///   </item>
/// </list>
/// <para>
/// A library that uses <c>ToOADate</c> on both the write and the read side round-trips its own
/// files perfectly and still hands Excel the wrong day for anything before March 1900. That is the
/// bug this type exists to prevent.
/// </para>
/// </summary>
internal static class ExcelSerialDate
{
    /// <summary>1 March 1900 — at and above this date the OLE Automation serial is already correct.</summary>
    private static readonly DateTime PhantomLeapDayEnd = new(1900, 3, 1);

    /// <summary>1 January 1900, Excel's serial 1 and the earliest date it can hold.</summary>
    public static readonly DateTime MinValue = new(1900, 1, 1);

    /// <summary>The serial Excel would show for <paramref name="value"/>, including its time of day.</summary>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The date precedes 1 January 1900, which Excel has no serial for.
    /// </exception>
    public static double ToSerial(DateTime value)
    {
        if (value < MinValue)
        {
            throw new ArgumentOutOfRangeException(
                nameof(value),
                value,
                $"Excel cannot represent a date before {MinValue:yyyy-MM-dd}; its serial numbering starts there. "
                + "Write the value as text if the workbook has to carry it.");
        }

        var serial = value.ToOADate();
        return value < PhantomLeapDayEnd ? serial - 1 : serial;
    }

    /// <summary>The date a SpreadsheetML <paramref name="serial"/> denotes.</summary>
    /// <remarks>
    /// Permissive by design, because this also runs against files this library did not write.
    /// Serial 60 is the phantom 29 February 1900 and has no <see cref="DateTime"/>; it comes back
    /// as 1 March 1900, the next day that does exist.
    /// </remarks>
    public static DateTime FromSerial(double serial) =>
        DateTime.FromOADate(serial < 61 ? serial + 1 : serial);
}
