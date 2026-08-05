using OfficeDocuments.Excel.PerformanceTests.Infrastructure;
using Xunit.Abstractions;

namespace OfficeDocuments.Excel.PerformanceTests;

/// <summary>
/// The paths that are quadratic today. These guards do <b>not</b> claim the behaviour is
/// acceptable — they pin the shape it currently has so a change cannot quietly make it worse
/// while the work to fix it is still outstanding (EXCEL-005).
/// <para>
/// The ceiling comes from the gap between complexity classes, not from a stopwatch. Quadrupling
/// the input costs roughly 4x for linear work, 16x for quadratic and 64x for cubic, so a ceiling
/// in the twenties clears today's quadratic curve with room for runner noise while still failing
/// if a path slips a class. When one of these is fixed, its test moves to
/// <see cref="LinearScalingGuards"/>; that migration is the definition of done.
/// </para>
/// <para>
/// Base sizes are chosen so the smaller measurement clears the noise floor in
/// <see cref="Measure"/>. They are smaller than the equivalent benchmark parameters because
/// <see cref="Measure"/> reports the fastest of several warm runs where BenchmarkDotNet reports
/// the mean of cold ones — the same workload measures two to four times faster here.
/// </para>
/// <para>
/// Three of the four hot spots are here. The fourth, distinct-style allocation, is guarded by
/// <see cref="AllocationGuards.DistinctStyleAllocation_GrowthStaysQuadratic"/> instead: it
/// allocates so heavily that garbage collection adds more to the wall clock than the difference
/// between a quadratic and a cubic curve, which makes a timing ratio unable to tell them apart.
/// Its allocation, measured directly, is a clean 4x for 2x the input.
/// </para>
/// </summary>
public class KnownHotSpotGuards(ITestOutputHelper output) : PerformanceGuard(output)
{
    /// <summary>Comfortably above quadratic (16), far below cubic (64).</summary>
    private const double QuadraticCeiling = 26.0;

    private const int Factor = 4;

    /// <summary>
    /// Each <c>SetComment</c> re-serializes the whole comments part and rebuilds the entire legacy
    /// VML drawing, so per-call cost grows with the number of comments already attached.
    /// <para>
    /// Benchmarked: 50 comments 10 ms / 5 MB, 100 comments 34 ms / 16 MB, 200 comments 166 ms /
    /// 59 MB — a clean 16x for 4x the input. The steepest of the four.
    /// </para>
    /// </summary>
    [Fact]
    public void CommentWriting_StaysWithinTheKnownQuadraticCeiling()
    {
        var ratio = Measure.GrowthRatio(Workloads.Comments, baseSize: 100, Factor);

        AssertGrowth(
            "SetComment, 4x as many comments", ratio, QuadraticCeiling,
            "The known cause is the full re-serialization of the comments part and the VML "
            + "drawing on every call in CommentWriter.Set.");
    }

    /// <summary>
    /// Writing one cell at a far column backfills every cell before it, and each backfilled cell
    /// is positioned by a linear scan.
    /// <para>
    /// Benchmarked: column 1 000 takes 11 ms, column 4 000 takes 101 ms — about 9x for 4x, so in
    /// this range the curve sits between linear and quadratic.
    /// </para>
    /// </summary>
    [Fact]
    public void FarColumnBackfill_StaysWithinTheKnownQuadraticCeiling()
    {
        var ratio = Measure.GrowthRatio(Workloads.CellAtFarColumn, baseSize: 2_000, Factor);

        AssertGrowth(
            "one cell, 4x further along the row", ratio, QuadraticCeiling,
            "The known cause is the linear insertion scan inside the backfill loop in "
            + "Row.CreateCell.");
    }

    /// <summary>
    /// Sorting deep-clones every cell in the range into a snapshot and then writes all of them
    /// back through the normal cell API.
    /// <para>
    /// The wall-clock cost of the sort is dominated by building the range it sorts, which makes a
    /// timing ratio here mostly a measurement of something else. The precise, stable statement
    /// about this path is the allocation one — see
    /// <see cref="AllocationGuards.RangeSort_AllocatesOneExtraCopyOfTheRange"/>. This guard only
    /// catches a change of complexity class.
    /// </para>
    /// </summary>
    [Fact]
    public void RangeSort_StaysWithinTheKnownQuadraticCeiling()
    {
        var ratio = Measure.GrowthRatio(Workloads.BuildAndSort, baseSize: 250, Factor);

        AssertGrowth(
            "build + SortByColumn, 4x as many rows", ratio, QuadraticCeiling,
            "The known cause is the clone-and-replay in Range.SortByColumn, not the comparisons.");
    }
}
