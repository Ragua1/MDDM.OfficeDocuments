using System.Diagnostics;

namespace OfficeDocuments.Excel.PerformanceTests.Infrastructure;

/// <summary>
/// The measurement primitives every guard in this project is built from.
/// <para>
/// Two rules shape all of it. First, <b>never assert on an absolute duration</b> — that measures
/// the machine, and a threshold tuned on a workstation is either useless or flaky on a shared CI
/// runner. Assert on the ratio between t(N) and t(4N) instead: it is computed from two
/// measurements taken seconds apart on the same hardware, so the hardware cancels out and what
/// is left is the algorithm.
/// </para>
/// <para>
/// Second, <b>take the fastest run, not the average</b>. Noise on a shared machine is one-sided:
/// a scheduler preemption or a background process can only make a run slower, never faster. The
/// minimum is therefore the closest thing to a clean measurement, and it is far more stable
/// across repeats than a mean that a single outlier can drag anywhere.
/// </para>
/// </summary>
/// <summary>
/// A growth measurement: the ratio plus the two timings it came from.
/// </summary>
internal readonly record struct Growth(double Ratio, double BaselineMs, double ScaledMs)
{
    public override string ToString() => $"{Ratio:F1}x  ({BaselineMs:F1} ms -> {ScaledMs:F1} ms)";
}

internal static class Measure
{
    /// <summary>
    /// Below this, a measurement is mostly JIT residue, GC timing and scheduler jitter, and a
    /// ratio computed from it says nothing. Guards fail loudly instead of passing by luck.
    /// </summary>
    private const double MeasurementFloorMs = 4;

    /// <summary>
    /// Five rather than three. Taking the minimum only helps if one of the runs happened to land
    /// in a quiet window, and five draws make that much likelier at a cost of seconds.
    /// </summary>
    private const int DefaultRepeats = 5;

    /// <summary>
    /// Runs <paramref name="work"/> once to warm up, then <paramref name="repeats"/> times, and
    /// returns the fastest elapsed time in milliseconds.
    /// </summary>
    public static double FastestMs(Action work, int repeats = DefaultRepeats)
    {
        // Warm-up. The first call pays for JIT, static initialization and first-touch page
        // faults, none of which the guard is trying to measure.
        work();

        var best = double.MaxValue;
        for (var i = 0; i < repeats; i++)
        {
            Settle();

            var start = Stopwatch.GetTimestamp();
            work();
            best = Math.Min(best, Stopwatch.GetElapsedTime(start).TotalMilliseconds);
        }

        return best;
    }

    /// <summary>
    /// How much more expensive <paramref name="work"/> gets when its input grows by
    /// <paramref name="factor"/>. Linear work returns roughly <paramref name="factor"/>;
    /// quadratic work returns roughly its square.
    /// <para>
    /// Both absolute timings come back with the ratio. A ratio on its own is unreadable when it
    /// looks wrong: 4x could mean the code is linear or it could mean the baseline measurement
    /// was inflated, and only the raw numbers distinguish those.
    /// </para>
    /// </summary>
    // `factor` has no default on purpose: a ratio means nothing without knowing what growth
    // produced it, so every call site states it next to the ceiling it is compared against.
    public static Growth GrowthRatio(Action<int> work, int baseSize, int factor, int repeats = DefaultRepeats) =>
        Compare(() => work(baseSize), () => work(baseSize * factor), baseSize, repeats);

    /// <summary>
    /// The same comparison for workloads whose two sizes cannot be expressed as one function of
    /// N — typically because the input has to be prepared outside the measured region.
    /// </summary>
    public static Growth Compare(Action baseline, Action scaled, int baseSize, int repeats = DefaultRepeats)
    {
        var small = FastestMs(baseline, repeats);
        var large = FastestMs(scaled, repeats);

        Assert.True(
            small >= MeasurementFloorMs,
            $"The baseline run took {small:F2} ms, below the {MeasurementFloorMs} ms floor, so the "
            + $"ratio would be noise. Raise the base size above {baseSize} in this test.");

        return new Growth(large / small, small, large);
    }

    /// <summary>
    /// Bytes allocated by one run of <paramref name="work"/> on the calling thread.
    /// <para>
    /// This is the guard worth having where it applies: allocation is counted, not sampled, so it
    /// does not flake, does not care what else the machine is doing, and is reproducible to the
    /// byte on a given runtime.
    /// </para>
    /// </summary>
    public static long AllocatedBytes(Action work)
    {
        // Same warm-up reasoning: one-time allocations behind the first call belong to neither
        // the measurement nor the budget the caller is asserting on.
        work();
        Settle();

        var before = GC.GetAllocatedBytesForCurrentThread();
        work();
        return GC.GetAllocatedBytesForCurrentThread() - before;
    }

    /// <summary>
    /// Puts the heap in a comparable state before each measurement, so one run does not pay for
    /// the garbage the previous one left behind.
    /// </summary>
    private static void Settle()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }
}
