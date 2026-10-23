namespace VB_build
{
    /// <summary>
    /// Abstraction over "run one VB6 build". Lets hosts (Web API job queue, tests) depend on an
    /// interface instead of the concrete <see cref="Controller"/>/<see cref="VbBuildRunner"/>,
    /// so tests can substitute a fake runner without needing VB6/TLI installed.
    /// </summary>
    public interface IBuildRunner
    {
        BuildResult Run(BuildRequest request);
    }
}
