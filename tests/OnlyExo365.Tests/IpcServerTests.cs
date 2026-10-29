using System.Collections.Concurrent;
using System.Text;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Worker.Ipc;

namespace OnlyExo365.Tests;

public class IpcServerTests
{
    [Fact]
    public async Task WriteSerializedMessageAsync_SerializesConcurrentWrites()
    {
        using var writeLock = new SemaphoreSlim(1, 1);
        using var writer = new ConcurrencyTrackingWriter();

        var tasks = Enumerable.Range(0, 24)
            .Select(index => IpcServer.WriteSerializedMessageAsync(writer, writeLock, $"message-{index}"));

        await Task.WhenAll(tasks);

        Assert.Equal(1, writer.MaxConcurrentWrites);
        Assert.Equal(24, writer.Messages.Count);
        Assert.Equal(24, writer.Messages.Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void CreatePendingRequestChannel_UsesConfiguredBoundedCapacity()
    {
        var channel = IpcServer.CreatePendingRequestChannel(1);

        Assert.True(channel.Writer.TryWrite(CreateRequest("corr-1")));
        Assert.False(channel.Writer.TryWrite(CreateRequest("corr-2")));
        Assert.Equal(1, channel.Reader.Count);
    }

    [Fact]
    public async Task CreatePendingRequestChannel_ReleasesWriterWhenReaderConsumesItem()
    {
        var channel = IpcServer.CreatePendingRequestChannel(1);

        await channel.Writer.WriteAsync(CreateRequest("corr-1"));
        var blockedWrite = channel.Writer.WriteAsync(CreateRequest("corr-2")).AsTask();

        Assert.False(blockedWrite.IsCompleted);

        var dequeued = await channel.Reader.ReadAsync();

        Assert.Equal("corr-1", dequeued.CorrelationId);
        await blockedWrite.WaitAsync(TimeSpan.FromSeconds(1));
        Assert.Equal(1, channel.Reader.Count);
    }

    [Fact]
    public void CanProcessConcurrently_IsReservedForUnifiedAuditLog()
    {
        Assert.True(IpcServer.CanProcessConcurrently(OperationType.SearchUnifiedAuditLog));
        Assert.False(IpcServer.CanProcessConcurrently(OperationType.GetComplianceWorkspace));
        Assert.False(IpcServer.CanProcessConcurrently(OperationType.GetMailboxes));
    }

    private static RequestEnvelope CreateRequest(string correlationId)
        => new()
        {
            CorrelationId = correlationId,
            Operation = OperationType.GetMailboxes
        };

    private sealed class ConcurrencyTrackingWriter : StringWriter
    {
        private int _currentWrites;
        private int _maxConcurrentWrites;

        public ConcurrentQueue<string> Messages { get; } = new();

        public int MaxConcurrentWrites => _maxConcurrentWrites;

        public override Encoding Encoding => Encoding.UTF8;

        public override async Task WriteLineAsync(ReadOnlyMemory<char> value, CancellationToken cancellationToken = default)
        {
            var activeWrites = Interlocked.Increment(ref _currentWrites);
            UpdateMax(activeWrites);

            try
            {
                await Task.Delay(15, cancellationToken).ConfigureAwait(false);
                var text = value.ToString();
                Messages.Enqueue(text);
                await base.WriteLineAsync(value, cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                Interlocked.Decrement(ref _currentWrites);
            }
        }

        public override async Task FlushAsync(CancellationToken cancellationToken)
        {
            await Task.Delay(5, cancellationToken).ConfigureAwait(false);
            await base.FlushAsync(cancellationToken).ConfigureAwait(false);
        }

        private void UpdateMax(int activeWrites)
        {
            while (true)
            {
                var snapshot = _maxConcurrentWrites;
                if (activeWrites <= snapshot)
                {
                    return;
                }

                if (Interlocked.CompareExchange(ref _maxConcurrentWrites, activeWrites, snapshot) == snapshot)
                {
                    return;
                }
            }
        }
    }
}

