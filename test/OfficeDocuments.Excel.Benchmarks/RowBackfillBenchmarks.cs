using BenchmarkDotNet.Attributes;

namespace OfficeDocuments.Excel.Benchmarks;

/// <summary>
/// Writing a single cell at a far column index. <c>Row.CreateCell</c> backfills every missing
/// cell up to the requested column so the children stay in ascending order, and each backfilled
/// cell is placed by <c>InsertCell</c>, which locates its slot with a linear scan. One cell at
/// column N therefore costs O(N²).
/// <para>
/// This is not a synthetic shape. A report that writes a wide header and then fills only a few
/// columns per row hits it on every row.
/// </para>
/// </summary>
[Config(typeof(ScalingConfig))]
public class RowBackfillBenchmarks
{
    [Params(1_000u, 2_000u, 4_000u)]
    public uint Column;

    /// <summary>One cell, at column N. Everything measured here is backfill.</summary>
    [Benchmark(Description = "single cell at far column")]
    public uint FarColumn()
    {
        using var stream = new MemoryStream();
        using var spreadsheet = Spreadsheet.CreateDocument(stream);
        var worksheet = spreadsheet.AddWorksheet("Sheet");
        var row = worksheet.AddRow(1);

        row.AddCell(Column, "value");

        return Column;
    }

    /// <summary>
    /// The same N cells written left to right. Each write extends the contiguous prefix by one,
    /// so no backfill happens — but <c>InsertCell</c> still scans. The gap between this and the
    /// benchmark above is the cost of the backfill specifically.
    /// </summary>
    [Benchmark(Description = "N cells written in order")]
    public uint Sequential()
    {
        using var stream = new MemoryStream();
        using var spreadsheet = Spreadsheet.CreateDocument(stream);
        var worksheet = spreadsheet.AddWorksheet("Sheet");
        var row = worksheet.AddRow(1);

        for (var column = 1u; column <= Column; column++)
        {
            row.AddCell(column, "value");
        }

        return Column;
    }
}
