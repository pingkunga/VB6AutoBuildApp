using System.Text;
using System.Threading;

namespace VB_build.Api.Services
{
    /// <summary>
    /// A <see cref="TextWriter"/> installed via <c>Console.SetOut</c> that transparently passes
    /// everything through to the original console output, while also capturing per-job progress
    /// lines for whichever <see cref="BuildJob"/> is "active" on the current async flow.
    ///
    /// <para>
    /// This lets <c>VB_build.Core</c>'s <c>Controller</c> keep writing plain
    /// <c>Console.Write</c>/<c>Console.WriteLine</c> calls (unchanged, same behavior as the
    /// console app) while the Web API observes those same lines as build progress, without any
    /// coupling between the two projects.
    /// </para>
    ///
    /// <para>
    /// <see cref="CurrentJob"/> is an <see cref="AsyncLocal{T}"/>, so it only captures output
    /// produced within the async call stack that set it (the background build worker while it is
    /// running a specific job); concurrent HTTP request handling/logging on other logical call
    /// contexts is unaffected.
    /// </para>
    /// </summary>
    public class JobConsoleWriter : TextWriter
    {
        /// <summary>The build job whose console output should be captured on this async flow, if any.</summary>
        public static readonly AsyncLocal<BuildJob?> CurrentJob = new();

        private readonly TextWriter inner;
        private readonly object bufferLock = new();
        private readonly StringBuilder buffer = new();

        public JobConsoleWriter(TextWriter inner)
        {
            this.inner = inner;
        }

        public override Encoding Encoding => inner.Encoding;

        public override void Write(char value)
        {
            inner.Write(value);
            AppendToBufferIfJobActive(value.ToString());
        }

        public override void Write(string? value)
        {
            if (value is null)
            {
                return;
            }

            inner.Write(value);
            AppendToBufferIfJobActive(value);
        }

        public override void WriteLine(string? value)
        {
            inner.WriteLine(value);
            FlushBufferedLine(value);
        }

        public override void WriteLine()
        {
            inner.WriteLine();
            FlushBufferedLine(null);
        }

        // Only the single background build worker thread ever writes while a job is active (VB6
        // builds run one at a time, see BuildQueueService), so the buffer is never touched
        // concurrently by unrelated writers (e.g. ASP.NET Core's own console logger thread) —
        // those bypass the buffer entirely, since CurrentJob.Value is null on their call context.
        private void AppendToBufferIfJobActive(string text)
        {
            if (CurrentJob.Value is null)
            {
                return;
            }

            lock (bufferLock)
            {
                buffer.Append(text);
            }
        }

        private void FlushBufferedLine(string? trailingValue)
        {
            BuildJob? job = CurrentJob.Value;
            if (job is null)
            {
                return;
            }

            string line;
            lock (bufferLock)
            {
                if (!string.IsNullOrEmpty(trailingValue))
                {
                    buffer.Append(trailingValue);
                }

                line = buffer.ToString();
                buffer.Clear();
            }

            if (line.Length > 0)
            {
                job.AppendLogLine(line);
            }
        }
    }
}
