using BenchmarkDotNet.Attributes;

namespace OfficeDocuments.Excel.Benchmarks;

/// <summary>
/// Attaching comments to cells. <c>CommentWriter.Set</c> does three things per call that each
/// depend on how many comments already exist: it scans the comment list for the reference,
/// serializes the whole comments part with <c>Comments.Save()</c>, and rebuilds the entire legacy
/// VML drawing from scratch — every shape, not just the new one.
/// <para>
/// Comments are therefore the steepest of the known hot spots: the per-call work is not a cheap
/// list walk but two full serializations of everything written so far.
/// </para>
/// </summary>
[Config(typeof(ScalingConfig))]
public class CommentBenchmarks
{
    [Params(50, 100, 200)]
    public int Count;

    [Benchmark(Description = "SetComment on N cells")]
    public int SetComments()
    {
        using var stream = new MemoryStream();
        using var spreadsheet = Spreadsheet.CreateDocument(stream);
        var worksheet = spreadsheet.AddWorksheet("Sheet");

        for (var i = 1; i <= Count; i++)
        {
            worksheet.AddCell(1u, (uint)i, $"value {i}")
                .SetComment($"comment {i}", "bench");
        }

        return Count;
    }

    /// <summary>
    /// The same N cells without comments. Subtracting this isolates the comment machinery from
    /// the cost of simply having N rows.
    /// </summary>
    [Benchmark(Description = "N cells, no comments", Baseline = true)]
    public int PlainCells()
    {
        using var stream = new MemoryStream();
        using var spreadsheet = Spreadsheet.CreateDocument(stream);
        var worksheet = spreadsheet.AddWorksheet("Sheet");

        for (var i = 1; i <= Count; i++)
        {
            worksheet.AddCell(1u, (uint)i, $"value {i}");
        }

        return Count;
    }
}
