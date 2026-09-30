using VB_build;

namespace VB_build.Api.Services
{
    /// <summary>
    /// Background worker that drains <see cref="BuildJobQueue"/> and runs each build job
    /// sequentially via <see cref="IBuildRunner"/>, updating <see cref="BuildJobStore"/> as it
    /// goes. Concurrency is intentionally fixed at 1 (see <see cref="BuildJobQueue"/>).
    /// </summary>
    public class BuildQueueService : BackgroundService
    {
        private readonly BuildJobQueue queue;
        private readonly BuildJobStore store;
        private readonly IBuildRunner runner;
        private readonly ILogger<BuildQueueService> logger;

        public BuildQueueService(BuildJobQueue queue, BuildJobStore store, IBuildRunner runner, ILogger<BuildQueueService> logger)
        {
            this.queue = queue;
            this.store = store;
            this.runner = runner;
            this.logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            await foreach (string jobId in queue.ReadAllAsync(stoppingToken))
            {
                if (!store.TryGet(jobId, out BuildJob job))
                {
                    continue;
                }

                job.Status = BuildJobStatus.Running;

                try
                {
                    BuildResult result = runner.Run(job.Request);
                    job.Result = result;
                    job.Status = result.Success ? BuildJobStatus.Completed : BuildJobStatus.Failed;
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, "Build job {JobId} threw an unhandled exception", jobId);
                    job.Result = new BuildResult { ErrorCount = 1, ErrorMessage = ex.Message };
                    job.Status = BuildJobStatus.Failed;
                }
                finally
                {
                    job.CompletedAt = DateTimeOffset.UtcNow;
                }
            }
        }
    }
}
