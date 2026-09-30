namespace VB_build.Api.Models
{
    /// <summary>Response body for both the trigger and status endpoints.</summary>
    public class BuildJobStatusResponse
    {
        public string JobId { get; set; } = string.Empty;

        /// <summary>"Queued", "Running", "Completed", or "Failed".</summary>
        public string Status { get; set; } = string.Empty;

        public bool? Success { get; set; }

        public int? ErrorCount { get; set; }

        public string? ErrorMessage { get; set; }

        public DateTimeOffset CreatedAt { get; set; }

        public DateTimeOffset? CompletedAt { get; set; }
    }
}
