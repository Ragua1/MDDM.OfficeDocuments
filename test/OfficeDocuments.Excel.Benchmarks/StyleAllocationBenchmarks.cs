using BenchmarkDotNet.Attributes;
using OfficeDocuments.Excel.Styles;

namespace OfficeDocuments.Excel.Benchmarks;

/// <summary>
/// The style dedup path. <c>Style.GetFontId</c> and its siblings look a candidate up by walking
/// every entry already in the stylesheet and comparing it structurally, so allocating N distinct
/// styles performs O(N²) element comparisons.
/// <para>
/// Two benchmarks, because the two directions behave completely differently and only one of them
/// is a problem: a report that reuses a handful of styles across thousands of cells keeps the
/// scanned list short, while a report that derives a style per row grows it without bound.
/// </para>
/// </summary>
[Config(typeof(ScalingConfig))]
public class StyleAllocationBenchmarks
{
    [Params(250, 500, 1_000)]
    public int Count;

    /// <summary>
    /// Every style is new, so every lookup misses and scans the whole list before appending.
    /// This is the quadratic case.
    /// </summary>
    [Benchmark(Description = "CreateStyle, all distinct")]
    public int Distinct()
    {
        using var stream = new MemoryStream();
        using var spreadsheet = Spreadsheet.CreateDocument(stream);

        for (var i = 0; i < Count; i++)
        {
            spreadsheet.CreateStyle(new Font { FontSize = 6 + i * 0.25 });
        }

        return Count;
    }

    /// <summary>
    /// The same call count against a fixed set of eight styles. The scanned list stays tiny, so
    /// this should stay linear — it is the control that shows the cost above comes from list
    /// growth and not from <c>CreateStyle</c> being slow per call.
    /// </summary>
    [Benchmark(Description = "CreateStyle, 8 reused")]
    public int Reused()
    {
        using var stream = new MemoryStream();
        using var spreadsheet = Spreadsheet.CreateDocument(stream);

        for (var i = 0; i < Count; i++)
        {
            spreadsheet.CreateStyle(new Font { FontSize = 8 + i % 8 });
        }

        return Count;
    }
}
