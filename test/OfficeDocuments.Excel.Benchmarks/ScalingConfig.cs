using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Loggers;

namespace OfficeDocuments.Excel.Benchmarks;

/// <summary>
/// The shared job for every benchmark in this assembly.
/// <para>
/// These workloads run for tens of milliseconds up to seconds, so the interesting number is
/// how the cost scales with N, not the nanosecond-level precision BenchmarkDotNet defaults to
/// chasing. One invocation per iteration and a low iteration count keep the whole suite to a
/// few minutes while still separating a quadratic curve from a linear one by an obvious margin.
/// </para>
/// </summary>
public sealed class ScalingConfig : ManualConfig
{
    public ScalingConfig()
    {
        AddJob(Job.Default
            .WithWarmupCount(1)
            .WithIterationCount(5)
            // Each workload builds and discards a whole workbook, so it must not be batched:
            // an unrolled loop would measure several workbooks as one operation.
            .WithInvocationCount(1)
            .WithUnrollFactor(1));

        AddColumnProvider(DefaultColumnProviders.Instance);
        AddColumn(RankColumn.Arabic);
        AddDiagnoser(MemoryDiagnoser.Default);
        AddLogger(ConsoleLogger.Default);
    }
}
