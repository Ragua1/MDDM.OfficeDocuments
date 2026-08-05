using BenchmarkDotNet.Running;

// BenchmarkSwitcher rather than BenchmarkRunner: the suite covers several independent hot spots
// and running all of them takes minutes, so the normal use is `-- --filter *StyleAllocation*`.
BenchmarkSwitcher
    .FromAssembly(typeof(Program).Assembly)
    .Run(args);

/// <summary>
/// Entry point marker. Top-level statements need a named type for <c>FromAssembly</c>.
/// </summary>
public partial class Program;
