using VB_build;

namespace VB_build.Api.Services
{
    public enum BuildJobStatus
    {
        Queued,
        Running,
        Completed,
        Failed
    }

    /// <summary>In-memory record of one queued/running/finished build job.</summary>
    public class BuildJob
    {
        public required string Id { get; init; }

        public required BuildRequest Request { get; init; }

        public BuildJobStatus Status { get; set; } = BuildJobStatus.Queued;

        public BuildResult? Result { get; set; }

        public DateTimeOffset CreatedAt { get; init; } = DateTimeOffset.UtcNow;

        public DateTimeOffset? CompletedAt { get; set; }
    }
}
