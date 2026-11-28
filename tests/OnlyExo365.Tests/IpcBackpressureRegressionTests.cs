using System.Reflection;
using System.Text;
using OnlyExo365.Contracts;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Worker.Ipc;

namespace OnlyExo365.Tests;

public sealed class IpcBackpressureRegressionTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BlockedServerSend_HasFiniteBudgetAndAbortsTransport(bool eventPipe)
    {
        using var fixture = new LocalPowerShellFixture();
        using var cancellation = new CancellationTokenSource();
        using var buffer = new MemoryStream();
        using var writer = new StreamWriter(buffer, Encoding.UTF8);
        using var server = new IpcServer(fixture.Engine, new IpcSessionContext { UserScope = Guid.NewGuid().ToString("N"), SessionId = 0 }, "offline-token");
        SetField(server, "_serverCts", cancellation);
        var gate = (SemaphoreSlim)GetField(server, eventPipe ? "_eventWriteLock" : "_requestWriteLock");
        SetField(server, eventPipe ? "_eventWriter" : "_requestWriter", writer);
        await gate.WaitAsync();
        try
        {
            var method = typeof(IpcServer).GetMethod(eventPipe ? "SendEventAsync" : "SendResponseRawAsync", BindingFlags.NonPublic | BindingFlags.Instance)!;
            var task = eventPipe
                ? (Task)method.Invoke(server, [new EventEnvelope { CorrelationId = "probe" }])!
                : (Task)method.MakeGenericMethod(typeof(HandshakeResponse)).Invoke(server, [new HandshakeResponse(), null])!;
            await task.WaitAsync(TimeSpan.FromSeconds(9));
            Assert.True(cancellation.IsCancellationRequested);
        }
        finally { gate.Release(); }
    }

    [Fact]
    public async Task ConcurrentRequests_CannotBypassAdmissionLimit()
    {
        using var fixture = new LocalPowerShellFixture();
        using var server = new IpcServer(fixture.Engine, new IpcSessionContext { UserScope = Guid.NewGuid().ToString("N"), SessionId = 0 }, "offline-token");
        var channel = IpcServer.CreatePendingRequestChannel();
        SetField(server, "_pendingRequests", channel);
        using var output = new MemoryStream();
        using var writer = new StreamWriter(output, new UTF8Encoding(false), leaveOpen: true);
        SetField(server, "_requestWriter", writer);
        var enqueue = typeof(IpcServer).GetMethod("EnqueueRequestAsync", BindingFlags.NonPublic | BindingFlags.Instance)!;
        for (var i = 0; i < IpcConstants.MaxPendingRequests + 1; i++)
        {
            await (Task)enqueue.Invoke(server, [Request(i), CancellationToken.None])!;
            // A concurrent dispatch removes the queue item but retains its active lifetime.
            Assert.True(channel.Reader.TryRead(out _));
        }
        await (Task)enqueue.Invoke(server, [Request(100), CancellationToken.None])!;
        Assert.False(channel.Reader.TryRead(out _));
        var response = JsonMessageSerializer.Deserialize<ResponseEnvelope>(Encoding.UTF8.GetString(output.ToArray()).Trim());
        Assert.Equal(ErrorCode.ServiceUnavailable, response!.Error!.Code);
    }

    private static RequestEnvelope Request(int id) => new() { CorrelationId = $"offline-{id}", Operation = OperationType.SearchUnifiedAuditLog };
    private static object GetField(object target, string name) => target.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(target)!;
    private static void SetField(object target, string name, object value) => target.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Instance)!.SetValue(target, value);
}
