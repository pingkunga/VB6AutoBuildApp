using VB_build;
using VB_build.Api.Services;
using Xunit;

namespace VB_build.Api.Tests
{
    public class JobConsoleWriterTests
    {
        private static BuildJob NewJob() => new()
        {
            Id = Guid.NewGuid().ToString("N"),
            Request = new BuildRequest()
        };

        [Fact]
        public void WriteThenWriteLine_CapturesCombinedLineOnActiveJob()
        {
            var passthrough = new StringWriter();
            var writer = new JobConsoleWriter(passthrough);
            BuildJob job = NewJob();

            JobConsoleWriter.CurrentJob.Value = job;
            try
            {
                writer.Write("Delete Old Temp File: ");
                writer.WriteLine("Successful...");
            }
            finally
            {
                JobConsoleWriter.CurrentJob.Value = null;
            }

            Assert.Equal(new[] { "Delete Old Temp File: Successful..." }, job.GetLogLinesSince(0));
            Assert.Equal("Delete Old Temp File: Successful...", job.CurrentStep);
            Assert.Contains("Delete Old Temp File: Successful...", passthrough.ToString());
        }

        [Fact]
        public void WriteLineWithFullMessage_IsCapturedAsSingleLine()
        {
            var passthrough = new StringWriter();
            var writer = new JobConsoleWriter(passthrough);
            BuildJob job = NewJob();

            JobConsoleWriter.CurrentJob.Value = job;
            try
            {
                writer.WriteLine("CompatibleMode: Binary Compatible");
            }
            finally
            {
                JobConsoleWriter.CurrentJob.Value = null;
            }

            Assert.Equal(new[] { "CompatibleMode: Binary Compatible" }, job.GetLogLinesSince(0));
        }

        [Fact]
        public void NoActiveJob_DoesNotThrow_AndStillPassesThrough()
        {
            var passthrough = new StringWriter();
            var writer = new JobConsoleWriter(passthrough);

            JobConsoleWriter.CurrentJob.Value = null;
            writer.WriteLine("Some unrelated output");

            Assert.Contains("Some unrelated output", passthrough.ToString());
        }

        [Fact]
        public void GetLogLinesSince_ReturnsOnlyNewLines()
        {
            var passthrough = new StringWriter();
            var writer = new JobConsoleWriter(passthrough);
            BuildJob job = NewJob();

            JobConsoleWriter.CurrentJob.Value = job;
            try
            {
                writer.WriteLine("Step 1: Successful...");
                writer.WriteLine("Step 2: Successful...");
                writer.WriteLine("Step 3: Successful...");
            }
            finally
            {
                JobConsoleWriter.CurrentJob.Value = null;
            }

            Assert.Equal(3, job.LogLineCount);
            Assert.Equal(new[] { "Step 2: Successful...", "Step 3: Successful..." }, job.GetLogLinesSince(1));
            Assert.Empty(job.GetLogLinesSince(3));
        }
    }
}
