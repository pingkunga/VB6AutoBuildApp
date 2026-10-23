using Xunit;

// Console.Out (see JobConsoleWriter) and JobConsoleWriter.CurrentJob are process-global/static,
// and multiple WebApplicationFactory<Program> hosts in this assembly share the same process.
// Disable cross-collection parallelization so build-progress capture for one test's in-process
// host can't interleave with another test class's host/background worker.
[assembly: CollectionBehavior(DisableTestParallelization = true)]
