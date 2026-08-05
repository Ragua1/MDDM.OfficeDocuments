using OfficeDocuments.Excel.PerformanceTests.Infrastructure;
using Xunit.Abstractions;

namespace OfficeDocuments.Excel.PerformanceTests;

/// <summary>
/// Paths that do a linear amount of work and are <b>promised</b> to keep doing so. A failure here
/// is a real regression: something inside a loop started depending on how much the loop had
/// already done.
/// <para>
/// The ceiling is not 4. A linear algorithm does not measure linearly once it allocates enough to
/// provoke garbage collection — the larger run pays for Gen1 and Gen2 collections the smaller one
/// never triggers, which is worth roughly another 2x on these workloads. The ceiling therefore
/// sits between "linear plus GC" and the 16x a genuinely quadratic path would produce, and the
/// deterministic proof of linearity lives in <see cref="AllocationGuards"/>, where GC cannot
/// distort it.
/// </para>
/// <para>
/// Contrast with <see cref="KnownHotSpotGuards"/>, which pins behaviour that is already
/// quadratic. The two look alike and mean opposite things, which is why they are separate files.
/// </para>
/// </summary>
public class LinearScalingGuards(ITestOutputHelper output) : PerformanceGuard(output)
{
    private const double LinearCeiling = 10.0;

    private const int Factor = 4;

    /// <summary>
    /// The bulk import path — the one most callers actually use. Measured ≈ 7x, of which about 4x
    /// is the work and the rest is collecting the garbage it produces.
    /// </summary>
    [Fact]
    public void BulkRowWriting_ScalesLinearly()
    {
        var ratio = Measure.GrowthRatio(Workloads.BulkRows, baseSize: 500, Factor);

        AssertGrowth(
            "AddRows<T>, 4x the rows", ratio, LinearCeiling,
            "AddRows<T> must stay linear in the number of rows; a ratio near 16 means a per-row "
            + "operation started scanning everything written before it.");
    }

    /// <summary>
    /// Reusing a small set of styles across many cells. This is the documented way around the
    /// quadratic dedup scan pinned in <see cref="KnownHotSpotGuards"/>, so it has to keep working:
    /// the list being scanned stays short no matter how many times the caller asks.
    /// </summary>
    [Fact]
    public void ReusedStyleAllocation_ScalesLinearly()
    {
        var ratio = Measure.GrowthRatio(Workloads.ReusedStyles, baseSize: 2_000, Factor);

        AssertGrowth(
            "CreateStyle from a fixed set of 8, 4x the calls", ratio, LinearCeiling,
            "Style reuse is the recommended way around the distinct-style cost, so it must not "
            + "itself depend on the number of requests already made.");
    }

    /// <summary>
    /// Opening a finished workbook and reading every row — the only guard here on the read path.
    /// <para>
    /// The workbooks are built before the measurement starts, and deliberately so. Timing a write
    /// and a read together yields a number that describes neither: reading is the smaller share
    /// of a round trip, so a read path that turned quadratic would still leave the combined ratio
    /// below any ceiling loose enough not to flake, and the regression would pass. A ratio has to
    /// be taken over one path at a time to mean anything.
    /// </para>
    /// <para>
    /// The base size is larger than the other two guards use because opening a package costs tens
    /// of milliseconds whatever is inside it; at 500 rows that constant was most of the baseline
    /// and flattened the ratio towards 1.
    /// </para>
    /// </summary>
    [Fact]
    public void ReadingBack_ScalesLinearly()
    {
        const int baseSize = 2_000;

        var small = Workloads.WorkbookBytes(baseSize);
        var large = Workloads.WorkbookBytes(baseSize * Factor);

        var ratio = Measure.Compare(
            () => Workloads.ReadBack(small, baseSize),
            () => Workloads.ReadBack(large, baseSize * Factor),
            baseSize);

        AssertGrowth(
            "open + read every row, 4x the rows", ratio, LinearCeiling,
            "Cell lookup goes through a dictionary on both the row and the column axis, so "
            + "reading N rows must cost N. A ratio near 16 means one of them became a scan.");
    }
}
