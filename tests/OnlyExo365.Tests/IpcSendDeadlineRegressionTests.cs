using System.Collections;
using System.Diagnostics;
using System.Reflection;
using OnlyExo365.Contracts;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Shell.Ipc;

namespace OnlyExo365.Tests;

public class IpcSendDeadlineRegressionTests
{
    [Theory]
    [InlineData("lock", false)]
    [InlineData("write", false)]
    [InlineData("response", false)]
    [InlineData("lock", true)]
    [InlineData("write", true)]
    [InlineData("response", true)]
    public async Task BudgetCoversEverySendPhaseAndCleansRegistrations(string phase, bool manual)
    {
        await using var client = new IpcClient(new IpcSessionContext { UserScope = "offline", SessionId = 0 }, "fake-session-token");
        var stream = new ControlledStream(phase == "write");
        using var writer = new StreamWriter(stream, leaveOpen: true);
        Set(client, "_isConnected", true);
        Set(client, "_requestWriter", writer);
        var sendLock = (SemaphoreSlim)Get(client, "_sendLock");
        if (phase == "lock") await sendLock.WaitAsync();
        using var cts = new CancellationTokenSource();
        if (manual) cts.CancelAfter(80);
        var clock = Stopwatch.StartNew();
        try
        {
            var request = new RequestEnvelope { Operation = OperationType.GetConnectors, TimeoutMs = manual ? 5000 : 80 };
            var sending = client.SendRequestAsync(request, _ => { }, cts.Token);
            if (manual) await Assert.ThrowsAnyAsync<OperationCanceledException>(() => sending.WaitAsync(TimeSpan.FromSeconds(2)));
            else await Assert.ThrowsAsync<TimeoutException>(() => sending.WaitAsync(TimeSpan.FromSeconds(2)));
            Assert.True(clock.Elapsed < TimeSpan.FromSeconds(1), clock.Elapsed.ToString());
            Assert.Equal(phase != "write", client.IsConnected);
            foreach (var field in new[] { "_pendingRequests", "_eventHandlers", "_eventCounts" })
                Assert.Equal(0, (int)Get(client, field).GetType().GetProperty("Count")!.GetValue(Get(client, field))!);
        }
        finally
        {
            Set(client, "_isConnected", false);
            if (phase == "lock") sendLock.Release();
            stream.Unblock();
            // Allow the independently bounded cancellation notification to release the send lock.
            await sendLock.WaitAsync(TimeSpan.FromSeconds(3));
            sendLock.Release();
        }
    }

    private static object Get(object target, string field) => target.GetType().GetField(field, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(target)!;
    private static void Set(object target, string field, object value) => target.GetType().GetField(field, BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(target, value);
    private sealed class ControlledStream(bool block) : MemoryStream
    {
        private readonly TaskCompletionSource _gate = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public void Unblock() => _gate.TrySetResult();
        public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
        {
            if (block) await _gate.Task.WaitAsync(cancellationToken);
            await base.WriteAsync(buffer, cancellationToken);
        }
    }
}
