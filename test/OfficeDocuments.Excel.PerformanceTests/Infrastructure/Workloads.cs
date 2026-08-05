using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Interfaces;
using OfficeDocuments.Excel.Styles;

namespace OfficeDocuments.Excel.PerformanceTests.Infrastructure;

/// <summary>
/// The workloads the guards measure, kept in one place so a test body is only a threshold and a
/// reason. Each one builds and discards a complete workbook over a <see cref="MemoryStream"/>:
/// the constant setup cost is a couple of milliseconds, far below anything being asserted on, and
/// paying it inside the measurement is much safer than sharing mutable state between runs.
/// <para>
/// These mirror the benchmarks in <c>OfficeDocuments.Excel.Benchmarks</c>. When one changes, the
/// other should change with it, or the thresholds here stop being traceable to a measurement.
/// </para>
/// </summary>
internal static class Workloads
{
    public sealed record Record(int Id, string Name, decimal Amount, DateTime Date);

    /// <summary>Allocates <paramref name="count"/> styles that are all different from each other.</summary>
    public static void DistinctStyles(int count)
    {
        using var stream = new MemoryStream();
        using var spreadsheet = Spreadsheet.CreateDocument(stream);

        for (var i = 0; i < count; i++)
        {
            spreadsheet.CreateStyle(new Font { FontSize = 6 + i * 0.25 });
        }
    }

    /// <summary>Makes <paramref name="count"/> style requests that resolve to one of eight styles.</summary>
    public static void ReusedStyles(int count)
    {
        using var stream = new MemoryStream();
        using var spreadsheet = Spreadsheet.CreateDocument(stream);

        for (var i = 0; i < count; i++)
        {
            spreadsheet.CreateStyle(new Font { FontSize = 8 + i % 8 });
        }
    }

    /// <summary>Bulk-writes <paramref name="rows"/> records through the reflection-driven path.</summary>
    public static void BulkRows(int rows)
    {
        using var stream = new MemoryStream();
        using var spreadsheet = Spreadsheet.CreateDocument(stream);
        spreadsheet.AddWorksheet("Sheet").AddRows(Records(rows), includeHeader: true);
    }

    /// <summary>Writes <paramref name="rows"/> records, closes the package, reopens and reads it back.</summary>
    public static int RoundTrip(int rows) => ReadBack(WorkbookBytes(rows), rows);

    /// <summary>
    /// A finished <c>.xlsx</c> holding <paramref name="rows"/> records, as bytes.
    /// <para>
    /// Separate from <see cref="ReadBack"/> so a guard can prepare its input outside the measured
    /// region. Timing a write and a read together produces a number that means neither: if the
    /// read side turned quadratic it would still be a minority of the total, and the combined
    /// ratio would land below any ceiling loose enough not to flake.
    /// </para>
    /// </summary>
    public static byte[] WorkbookBytes(int rows)
    {
        using var stream = new MemoryStream();

        using (var spreadsheet = Spreadsheet.CreateDocument(stream))
        {
            spreadsheet.AddWorksheet("Sheet").AddRows(Records(rows), includeHeader: true);
            spreadsheet.Close();
        }

        return stream.ToArray();
    }

    /// <summary>Opens a finished workbook and reads one column out of every row.</summary>
    public static int ReadBack(byte[] workbook, int rows)
    {
        // A MemoryStream over a byte[] is fixed-size, and the library opens streams for editing
        // by default; copy into an expandable one even though this open is read-only.
        using var stream = new MemoryStream(workbook.Length);
        stream.Write(workbook, 0, workbook.Length);
        stream.Position = 0;

        using var reopened = Spreadsheet.OpenDocument(stream, isEditable: false);
        var worksheet = reopened.GetWorksheet("Sheet")
                        ?? throw new InvalidOperationException("The reopened workbook lost its worksheet.");

        var read = 0;
        for (var rowIndex = 2u; rowIndex <= rows + 1; rowIndex++)
        {
            if (!string.IsNullOrEmpty(worksheet.GetCell(2u, rowIndex)?.GetStringValue()))
            {
                read++;
            }
        }

        return read;
    }

    /// <summary>Attaches a comment to each of <paramref name="count"/> cells.</summary>
    public static void Comments(int count)
    {
        using var stream = new MemoryStream();
        using var spreadsheet = Spreadsheet.CreateDocument(stream);
        var worksheet = spreadsheet.AddWorksheet("Sheet");

        for (var i = 1; i <= count; i++)
        {
            worksheet.AddCell(1u, (uint)i, $"value {i}").SetComment($"comment {i}", "guard");
        }
    }

    /// <summary>Writes one cell at <paramref name="column"/>, forcing a backfill of everything before it.</summary>
    public static void CellAtFarColumn(int column)
    {
        using var stream = new MemoryStream();
        using var spreadsheet = Spreadsheet.CreateDocument(stream);
        spreadsheet.AddWorksheet("Sheet").AddRow(1).AddCell((uint)column, "value");
    }

    /// <summary>Builds a <paramref name="rows"/>-row block in descending order, then sorts it ascending.</summary>
    public static void BuildAndSort(int rows)
    {
        const uint columns = 5;

        using var stream = new MemoryStream();
        using var spreadsheet = Spreadsheet.CreateDocument(stream);
        var worksheet = Populate(spreadsheet, rows, columns);

        worksheet.GetRange(1, 1, columns, (uint)rows).SortByColumn(1, SortDirection.Ascending);
    }

    /// <summary>The same block, built but not sorted — the baseline the sort guard subtracts.</summary>
    public static void BuildOnly(int rows)
    {
        using var stream = new MemoryStream();
        using var spreadsheet = Spreadsheet.CreateDocument(stream);
        Populate(spreadsheet, rows, 5);
    }

    private static IWorksheet Populate(ISpreadsheet spreadsheet, int rows, uint columns)
    {
        var worksheet = spreadsheet.AddWorksheet("Sheet");

        for (var rowIndex = 1u; rowIndex <= rows; rowIndex++)
        {
            // Descending, so the sort has to move essentially every row.
            worksheet.AddCell(1u, rowIndex, rows - (int)rowIndex);
            for (var column = 2u; column <= columns; column++)
            {
                worksheet.AddCell(column, rowIndex, $"r{rowIndex}c{column}");
            }
        }

        return worksheet;
    }

    public static Record[] Records(int count) =>
        Enumerable.Range(1, count)
            .Select(i => new Record(i, $"Item {i}", i * 1.5m, new DateTime(2026, 1, 1).AddDays(i % 365)))
            .ToArray();
}
