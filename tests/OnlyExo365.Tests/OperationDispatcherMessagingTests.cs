using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Worker.Operations;

namespace OnlyExo365.Tests;

public class OperationDispatcherMessagingTests
{
    [Fact]
    public async Task SendProgressAsync_PublishesProgressEnvelopeWithCounters()
    {
        EventEnvelope? captured = null;
        var publisher = new OperationEventPublisher(evt =>
        {
            captured = evt;
            return Task.CompletedTask;
        });

        await publisher.SendProgressAsync("corr-progress", 65, "Working", 13, 20);

        Assert.NotNull(captured);
        Assert.Equal("corr-progress", captured!.CorrelationId);
        Assert.Equal(EventType.Progress, captured.EventType);

        var payload = JsonMessageSerializer.ExtractPayload<ProgressEventPayload>(captured.Payload);
        Assert.NotNull(payload);
        Assert.Equal(65, payload!.PercentComplete);
        Assert.Equal("Working", payload.StatusMessage);
        Assert.Equal(13, payload.CurrentItem);
        Assert.Equal(20, payload.TotalItems);
    }

    [Theory]
    [InlineData("info", LogLevel.Information)]
    [InlineData("warning", LogLevel.Warning)]
    [InlineData("DEBUG", LogLevel.Debug)]
    [InlineData("unknown", LogLevel.Information)]
    public void ParseLogLevel_NormalizesSupportedAliases(string rawLevel, LogLevel expected)
    {
        var publisher = new OperationEventPublisher(_ => Task.CompletedTask);

        var parsed = publisher.ParseLogLevel(rawLevel);

        Assert.Equal(expected, parsed);
    }

    [Fact]
    public void CreateResponses_PreserveResponseEnvelopeContract()
    {
        var factory = new OperationResponseFactory();

        var success = factory.CreateSuccess("corr-success", new SamplePayload { Value = "ok" });
        var error = factory.CreateError("corr-error", ErrorCode.Timeout, "timeout", isTransient: true, retryAfterSeconds: 30);
        var cancelled = factory.CreateCancelled("corr-cancelled");

        Assert.True(success.Success);
        Assert.Equal("corr-success", success.CorrelationId);
        Assert.False(success.WasCancelled);
        Assert.NotNull(success.Payload);
        Assert.Equal("ok", JsonMessageSerializer.ExtractPayload<SamplePayload>(success.Payload)!.Value);

        Assert.False(error.Success);
        Assert.Equal("corr-error", error.CorrelationId);
        Assert.NotNull(error.Error);
        Assert.Equal(ErrorCode.Timeout, error.Error!.Code);
        Assert.True(error.Error.IsTransient);
        Assert.Equal(30, error.Error.RetryAfterSeconds);

        Assert.False(cancelled.Success);
        Assert.True(cancelled.WasCancelled);
        Assert.Equal("corr-cancelled", cancelled.CorrelationId);
        Assert.Null(cancelled.Error);
    }

    private sealed class SamplePayload
    {
        public string? Value { get; init; }
    }
}

