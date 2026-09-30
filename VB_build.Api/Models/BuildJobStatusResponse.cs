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

            /// <summary>
            /// The most recent build-progress line captured so far (e.g. "Build Project:
            /// Successful..."), or null if nothing has been captured yet. Always reflects the very
            /// latest step regardless of the "since" query parameter.
            /// </summary>
            public string? CurrentStep { get; set; }

            /// <summary>
            /// Build-progress lines with index &gt;= the "since" query parameter (default 0 = all
            /// lines so far), in order, e.g. "Delete Old Temp File: Successful...".
            /// </summary>
            public string[] LogLines { get; set; } = Array.Empty<string>();

            /// <summary>
            /// Pass this value back as the "since" query parameter on the next poll to receive only
            /// new lines produced after this response was generated.
            /// </summary>
            public int NextSince { get; set; }
        }
}
