using OfficeDocuments.Excel.PerformanceTests.Infrastructure;
using Xunit.Abstractions;

namespace OfficeDocuments.Excel.PerformanceTests;

/// <summary>
/// Guards on allocated bytes rather than elapsed time.
/// <para>
/// These are the strongest tests in this project. Allocation is counted by the runtime, not
/// sampled by a clock: the number does not move when the machine is busy, needs no warm-up to be
/// trustworthy, and reproduces to within a rounding error on a given runtime. Where a defect
/// shows up in allocation at all, guard it here rather than with a stopwatch.
/// </para>
/// <para>
/// Figures in the comments were measured on a Ryzen 9 5900X / .NET 10 through
/// <c>OfficeDocuments.Excel.Benchmarks</c>. Ceilings sit well above them, so ordinary variation
/// between runtimes does not fail the build.
/// </para>
/// </summary>
public class AllocationGuards(ITestOutputHelper output) : PerformanceGuard(output)
{
    private const long Kilobyte = 1024;

    /// <summary>
    /// The bulk import path allocates a fixed amount per four-column row, flat across 2 000,
    /// 5 000 and 10 000 rows. Flat is the property being protected: a regression that made it
    /// grow with the row count would turn a large report from slow into impossible.
    /// <para>
    /// It measures about 12 KB here and about 19 KB under BenchmarkDotNet. The difference is not
    /// noise — this project disables tiered compilation, and unoptimized tier-0 code allocates
    /// more because the JIT has not yet proved which locals can stay off the heap. The ceiling
    /// clears both.
    /// </para>
    /// </summary>
    [Fact]
    public void BulkRowWriting_AllocatesABoundedAmountPerRow()
    {
        const int rows = 2_000;
        const long ceilingKb = 32; // measured 12 KB here, 19 KB under BenchmarkDotNet

        var perRowKb = Measure.AllocatedBytes(() => Workloads.BulkRows(rows)) / rows / Kilobyte;

        AssertKilobytes(
            "AddRows<T> per row", perRowKb, ceilingKb,
            "Either a new per-cell allocation appeared, or the per-row cost started depending on "
            + "how many rows came before it.");
    }

    /// <summary>
    /// Writing 4x the rows must allocate about 4x as much. This is the allocation-side twin of
    /// <see cref="LinearScalingGuards.BulkRowWriting_ScalesLinearly"/> and the one that states
    /// the linearity claim precisely — a stopwatch cannot, because garbage collection makes even
    /// perfectly linear work measure superlinearly once the heap gets busy.
    /// </summary>
    [Fact]
    public void BulkRowWriting_AllocationScalesLinearly()
    {
        var small = Measure.AllocatedBytes(() => Workloads.BulkRows(500));
        var large = Measure.AllocatedBytes(() => Workloads.BulkRows(2_000));

        AssertRatio(
            "AddRows<T> allocation, 4x the rows", (double)large / small, 5.0,
            "Allocation must track the amount of data written, not the square of it.");
    }

    /// <summary>
    /// <c>Range.SortByColumn</c> snapshots the range by deep-cloning every cell before writing any
    /// of them back, so a sort costs a second copy of the data. Measured at exactly 2.0x across
    /// 500, 1 000 and 2 000 rows.
    /// <para>
    /// Pinned, not endorsed: an in-place reorder would not need the copy at all. The ceiling
    /// exists so one clone cannot silently become two.
    /// </para>
    /// </summary>
    [Fact]
    public void RangeSort_AllocatesOneExtraCopyOfTheRange()
    {
        const int rows = 500;

        var build = Measure.AllocatedBytes(() => Workloads.BuildOnly(rows));
        var buildAndSort = Measure.AllocatedBytes(() => Workloads.BuildAndSort(rows));

        AssertRatio(
            "SortByColumn allocation vs building alone", (double)buildAndSort / build, 2.5,
            "Sorting is known to cost one full clone of the range; more than that means a second "
            + "copy crept in.");
    }

    /// <summary>
    /// Comments are the most expensive feature in the library per unit of content: the whole
    /// comments part and the whole VML drawing are rebuilt on every call. At 100 comments this
    /// comes to roughly 160 KB each, against about 7 KB for a plain cell.
    /// <para>
    /// The ceiling is per comment at a fixed N rather than a growth ratio, because the point here
    /// is the constant — the growth is already pinned in <see cref="KnownHotSpotGuards"/>.
    /// </para>
    /// </summary>
    [Fact]
    public void CommentWriting_StaysWithinItsKnownAllocationCost()
    {
        const int comments = 100;
        const long ceilingKb = 320; // measured ~160 KB

        var perCommentKb = Measure.AllocatedBytes(() => Workloads.Comments(comments)) / comments / Kilobyte;

        AssertKilobytes(
            "SetComment per comment", perCommentKb, ceilingKb,
            "The comments part and the VML drawing are rebuilt on every call; this pins how much "
            + "that costs until EXCEL-005 removes it.");
    }

    /// <summary>
    /// The complexity guard for the distinct-style path, which lives here rather than in
    /// <see cref="KnownHotSpotGuards"/> for a specific reason: it allocates so heavily — 300 MB
    /// for 500 styles, 1.2 GB for 1 000 — that garbage collection adds more to the wall clock
    /// than the difference between a quadratic and a cubic curve, so a timing ratio cannot tell
    /// those apart. It measures 33x for 4x the input where quadratic is 16x and cubic is 64x,
    /// which is a verdict of "somewhere in between" and therefore no verdict at all.
    /// <para>
    /// Allocation says it cleanly. Doubling the input allocates 4x as much: quadratic, exactly as
    /// the O(N²) comparison scan in <c>Style.GetFontId</c> predicts. A cubic regression would show
    /// as 8x and fail here immediately.
    /// </para>
    /// </summary>
    [Fact]
    public void DistinctStyleAllocation_GrowthStaysQuadratic()
    {
        var small = Measure.AllocatedBytes(() => Workloads.DistinctStyles(250));
        var large = Measure.AllocatedBytes(() => Workloads.DistinctStyles(500));

        AssertRatio(
            "CreateStyle allocation, 2x as many distinct styles", (double)large / small, 5.5,
            "The dedup scan is quadratic, which allocates 4x for 2x the input. Above roughly 5.5x "
            + "it has become something worse, not merely slow.");
    }

    /// <summary>
    /// The absolute figure behind the ratio above: 500 distinct styles cost about 300 MB. Pinned
    /// so the constant cannot grow either, before EXCEL-005 removes the scan entirely.
    /// </summary>
    [Fact]
    public void DistinctStyleAllocation_StaysWithinItsKnownCost()
    {
        const int styles = 500;
        const long ceilingKb = 512 * Kilobyte; // 512 MB; measured ~300 MB

        var allocatedKb = Measure.AllocatedBytes(() => Workloads.DistinctStyles(styles)) / Kilobyte;

        AssertKilobytes(
            $"CreateStyle, {styles} distinct styles", allocatedKb, ceilingKb,
            "The scan in Style.GetFontId compares against every existing entry; nothing should be "
            + "making that worse.");
    }
}
