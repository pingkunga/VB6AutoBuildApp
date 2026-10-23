using System.Collections.Concurrent;

namespace VB_build.Api.Services
{
    /// <summary>Thread-safe in-memory store of build jobs, keyed by job id.</summary>
    public class BuildJobStore
    {
        private readonly ConcurrentDictionary<string, BuildJob> jobs = new();

        public void Add(BuildJob job) => jobs[job.Id] = job;

        public bool TryGet(string jobId, out BuildJob job) => jobs.TryGetValue(jobId, out job!);
    }
}
