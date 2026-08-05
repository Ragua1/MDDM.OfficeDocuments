using OfficeDocuments.Excel.PerformanceTests.Infrastructure;

namespace OfficeDocuments.Excel.PerformanceTests;

/// <summary>
/// What the library is documented to handle, asserted rather than assumed.
/// <para>
/// These are not timing tests — they assert correctness at a size nothing else in the suite
/// reaches. The integration and verification tiers work with tens of rows, which is the right
/// choice for them but leaves a whole class of defect uncovered: index arithmetic that overflows,
/// a column reference that stops round-tripping past a certain width, a reopened workbook that
/// silently drops the tail. Those only appear at scale.
/// </para>
/// </summary>
public class ScaleCeilingTests
{
    /// <summary>
    /// The documented working size for a report. Every row must survive the write, the close and
    /// the reopen — the assertion is on the count that reads back, not on how long it took.
    /// </summary>
    [Fact]
    public void RoundTrip_At25000Rows_ReadsEveryRowBack()
    {
        const int rows = 25_000;

        var read = Workloads.RoundTrip(rows);

        Assert.Equal(rows, read);
    }

    /// <summary>
    /// Column 16 384 is the last one SpreadsheetML allows (<c>XFD</c>). Reaching it exercises the
    /// three-letter branch of the column-reference conversion and the backfill loop at its widest
    /// legal extent, and it is the point where an arithmetic mistake in either would show.
    /// </summary>
    [Fact]
    public void Worksheet_AcceptsTheLastLegalColumn()
    {
        const uint lastColumn = 16_384;

        using var stream = new MemoryStream();

        using (var spreadsheet = Spreadsheet.CreateDocument(stream))
        {
            var worksheet = spreadsheet.AddWorksheet("Wide");
            worksheet.AddRow(1).AddCell(lastColumn, "edge");
            spreadsheet.Close();
        }

        stream.Position = 0;

        using var reopened = Spreadsheet.OpenDocument(stream, isEditable: false);
        var reopenedSheet = reopened.GetWorksheet("Wide");

        Assert.Equal("edge", reopenedSheet?.GetCell(lastColumn, 1)?.GetStringValue());
    }
}
