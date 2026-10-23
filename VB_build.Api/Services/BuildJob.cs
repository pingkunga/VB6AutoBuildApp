using System.Collections.Generic;
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

        /// <summary>
        /// Step-by-step progress lines captured from the build's console output (see
        /// <see cref="JobConsoleWriter"/>), e.g. "Delete Old Temp File: Successful...",
        /// "Build Project: Successful...". Append-only, in order. Thread-safe: writer
        /// (background worker) and readers (status-polling requests) run concurrently.
        /// </summary>
        private readonly List<string> logLines = new();
        private readonly object logLock = new();

        public void AppendLogLine(string line)
        {
            lock (logLock)
            {
                logLines.Add(line);
            }
        }

        /// <summary>Snapshot of log lines with index &gt;= <paramref name="since"/>.</summary>
        public string[] GetLogLinesSince(int since)
        {
            lock (logLock)
            {
                if (since <= 0)
                {
                    return logLines.ToArray();
                }

                if (since >= logLines.Count)
                {
                    return Array.Empty<string>();
                }

                return logLines.GetRange(since, logLines.Count - since).ToArray();
            }
        }

        /// <summary>Total number of log lines captured so far.</summary>
        public int LogLineCount
        {
            get
            {
                lock (logLock)
                {
                    return logLines.Count;
                }
            }
        }

        /// <summary>The most recent log line, or null if nothing has been captured yet.</summary>
        public string? CurrentStep
        {
            get
            {
                lock (logLock)
                {
                    return logLines.Count == 0 ? null : logLines[^1];
                }
            }
        }
    }
}
