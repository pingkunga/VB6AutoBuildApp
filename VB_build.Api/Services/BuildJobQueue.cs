using System.Threading.Channels;

namespace VB_build.Api.Services
{
    /// <summary>
    /// Single-consumer FIFO queue of job ids waiting to be built. VB6 IDE COM automation (TLI)
    /// and VB6.EXE are not safe to run concurrently, so only one build ever executes at a time
    /// (see <see cref="BuildQueueService"/>).
    /// </summary>
    public class BuildJobQueue
    {
        private readonly Channel<string> channel = Channel.CreateUnbounded<string>();

        public void Enqueue(string jobId) => channel.Writer.TryWrite(jobId);

        public IAsyncEnumerable<string> ReadAllAsync(CancellationToken cancellationToken) =>
            channel.Reader.ReadAllAsync(cancellationToken);
    }
}
