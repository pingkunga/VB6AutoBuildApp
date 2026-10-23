using VB_build;

namespace VB_build.Api.Tests
{
    /// <summary>
    /// Test double for <see cref="IBuildRunner"/> so integration tests never invoke the real
    /// VB6.EXE/TLI build pipeline (which requires VB6 to be installed on the machine).
    /// </summary>
    public class FakeBuildRunner : IBuildRunner
    {
        public Func<BuildRequest, BuildResult> RunHandler { get; set; } =
            _ => new BuildResult { ErrorCount = 0, StepResults = new bool[10] };

        public List<BuildRequest> Received { get; } = new();

        public BuildResult Run(BuildRequest request)
        {
            Received.Add(request);
            return RunHandler(request);
        }
    }
}
