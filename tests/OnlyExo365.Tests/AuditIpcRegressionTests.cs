using OnlyExo365.Contracts;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Shell.Ipc;
using OnlyExo365.Worker.Ipc;

namespace OnlyExo365.Tests;

public sealed class AuditIpcRegressionTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CancellationWhileQueued_PreventsDispatch(bool deadline)
    {
        using var fixture = new LocalPowerShellFixture();
        await DefineSlowTraceAsync(fixture);
        using var server = CreateServer(fixture, out var context);
        var start = server.StartAsync();
        await using var client = new IpcClient(context, "dummy-test-token");
        await client.ConnectAsync();
        await start;
        try
        {
            var running = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var first = client.SendRequestAsync(TraceRequest(), evt => running.TrySetResult());
            await running.Task.WaitAsync(TimeSpan.FromSeconds(3));
            using var cancellation = new CancellationTokenSource(TimeSpan.FromMilliseconds(100));
            var request = TraceRequest();
            if (deadline) { request.TimeoutMs = 100; }
            var queued = client.SendRequestAsync(request, cancellationToken: deadline ? default : cancellation.Token);
            if (deadline) { await Assert.ThrowsAsync<TimeoutException>(async () => await queued); }
            else { await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await queued); }
            await first.WaitAsync(TimeSpan.FromSeconds(3));
            await Task.Delay(100);
            var count = await fixture.Engine.ExecuteAsync("$global:traceCalls");
            Assert.Equal(1, Convert.ToInt32(count.Output.Single().BaseObject));
        }
        finally
        {
            await server.StopAsync();
        }
    }

    [Fact]
    public async Task FullQueue_RejectsOverflowAndStillAnswersHeartbeat()
    {
        using var fixture = new LocalPowerShellFixture();
        await DefineSlowTraceAsync(fixture);
        using var server = CreateServer(fixture, out var context);
        var start = server.StartAsync();
        await using var client = new IpcClient(context, "dummy-test-token");
        await client.ConnectAsync();
        await start;
        using var cancellation = new CancellationTokenSource();
        var requests = new List<Task<ResponseEnvelope>>();
        try
        {
            var running = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            requests.Add(client.SendRequestAsync(TraceRequest(), evt => running.TrySetResult(), cancellation.Token));
            await running.Task.WaitAsync(TimeSpan.FromSeconds(3));
            for (var index = 0; index < IpcConstants.MaxPendingRequests; index++)
            {
                requests.Add(client.SendRequestAsync(TraceRequest(), cancellationToken: cancellation.Token));
            }
            var overflow = await client.SendRequestAsync(TraceRequest()).WaitAsync(TimeSpan.FromMilliseconds(500));
            Assert.False(overflow.Success);
            Assert.Equal(ErrorCode.ServiceUnavailable, overflow.Error?.Code);
            var heartbeat = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            client.HeartbeatReceived += (_, _) => heartbeat.TrySetResult();
            await client.SendHeartbeatAsync(42);
            await heartbeat.Task.WaitAsync(TimeSpan.FromMilliseconds(500));
        }
        finally
        {
            await cancellation.CancelAsync();
            foreach (var request in requests)
            {
                try { await request.WaitAsync(TimeSpan.FromSeconds(3)); }
                catch (OperationCanceledException) { }
            }
            await server.StopAsync();
        }
    }

    [Fact]
    public async Task RequestDeadline_IsTimeoutRatherThanUserCancellation()
    {
        using var fixture = new LocalPowerShellFixture();
        await DefineSlowTraceAsync(fixture);
        using var server = CreateServer(fixture, out var context);
        var start = server.StartAsync();
        await using var client = new IpcClient(context, "dummy-test-token");
        await client.ConnectAsync();
        await start;
        try
        {
            var request = TraceRequest();
            request.TimeoutMs = 100;
            try
            {
                var response = await client.SendRequestAsync(request);
                Assert.False(response.WasCancelled);
                Assert.Equal(ErrorCode.Timeout, response.Error?.Code);
            }
            catch (TimeoutException)
            {
                // The local deadline can win the race with the worker response.
            }
        }
        finally
        {
            await server.StopAsync();
        }
    }

    private static Task DefineSlowTraceAsync(LocalPowerShellFixture fixture) => fixture.DefineCommandsAsync("""
        $global:traceCalls = 0
        function Get-MessageTraceV2 { $global:traceCalls++; Start-Sleep -Milliseconds 1500 }
        """);

    private static IpcServer CreateServer(LocalPowerShellFixture fixture, out IpcSessionContext context)
    {
        context = new IpcSessionContext { UserScope = Guid.NewGuid().ToString("N"), SessionId = 0 };
        return new IpcServer(fixture.Engine, context, "dummy-test-token");
    }

    private static RequestEnvelope TraceRequest() => new()
    {
        Operation = OperationType.GetMessageTrace,
        TimeoutMs = 5000,
        Payload = JsonMessageSerializer.ToJsonElement(new GetMessageTraceRequest
        {
            StartDate = DateTime.UtcNow.AddDays(-2), EndDate = DateTime.UtcNow.AddDays(-1)
        })
    };
}
