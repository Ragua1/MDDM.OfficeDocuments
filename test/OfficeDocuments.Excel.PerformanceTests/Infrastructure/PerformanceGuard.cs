using Xunit.Abstractions;

namespace OfficeDocuments.Excel.PerformanceTests.Infrastructure;

/// <summary>
/// Base class for the guards. Its only job is to make every measurement visible, whether or not
/// the test passed.
/// <para>
/// A performance guard that only speaks when it fails is close to useless: the interesting signal
/// is drift, and drift is invisible until the day it crosses the threshold and someone has to
/// work out which of the last hundred commits caused it. Reporting the measured value on every
/// run turns the test log into a cheap trend line —
/// <c>dotnet test ... --logger "console;verbosity=detailed"</c> prints it.
/// </para>
/// </summary>
public abstract class PerformanceGuard(ITestOutputHelper output)
{
    /// <summary>
    /// Records a growth ratio and asserts it against <paramref name="ceiling"/>.
    /// </summary>
    protected void AssertRatio(string what, double measured, double ceiling, string diagnosis)
    {
        output.WriteLine($"{what}: {measured:F1}x  (ceiling {ceiling:F1}x)");
        Assert.True(measured <= ceiling, $"{what} measured {measured:F1}x against a {ceiling:F1}x ceiling. {diagnosis}");
    }

    /// <summary>
    /// Records a timed growth measurement, including the two durations behind the ratio, and
    /// asserts the ratio against <paramref name="ceiling"/>.
    /// </summary>
    internal void AssertGrowth(string what, Growth measured, double ceiling, string diagnosis)
    {
        output.WriteLine($"{what}: {measured}  (ceiling {ceiling:F1}x)");
        Assert.True(
            measured.Ratio <= ceiling,
            $"{what} measured {measured} against a {ceiling:F1}x ceiling. {diagnosis}");
    }

    /// <summary>
    /// Records an allocation figure in kilobytes and asserts it against <paramref name="ceilingKb"/>.
    /// </summary>
    protected void AssertKilobytes(string what, long measuredKb, long ceilingKb, string diagnosis)
    {
        output.WriteLine($"{what}: {measuredKb:N0} KB  (ceiling {ceilingKb:N0} KB)");
        Assert.True(measuredKb <= ceilingKb, $"{what} measured {measuredKb:N0} KB against a {ceilingKb:N0} KB ceiling. {diagnosis}");
    }
}
