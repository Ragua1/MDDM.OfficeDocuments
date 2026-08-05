// Timing measurements taken while other test classes run on other cores are not measurements.
// Every guard in this assembly compares two durations, so the whole assembly runs serially.
[assembly: CollectionBehavior(DisableTestParallelization = true)]
