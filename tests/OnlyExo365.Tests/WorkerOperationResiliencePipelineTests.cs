using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Errors;
using OnlyExo365.Shell.Resilience;
using OnlyExo365.Contracts.Results;
using OnlyExo365.Shell.Ipc;

namespace OnlyExo365.Tests;

public class WorkerOperationResiliencePipelineTests
{
    [Fact]
    public async Task ExecuteAsync_RetriesTransientFailureForIdempotentOperation()
    {
        var attempts = 0;
        var pipeline = CreatePipeline();

        var result = await pipeline.ExecuteAsync(
            OperationType.GetMailboxes,
            _ =>
            {
                attempts++;
                return Task.FromResult(attempts < 3
                    ? Result<string>.Failure(NormalizedError.Create(ErrorCode.Timeout, "timeout", isTransient: true))
                    : Result<string>.Success("ok"));
            },
            CancellationToken.None);

        Assert.True(result.IsSuccess);
        Assert.Equal("ok", result.Value);
        Assert.Equal(3, attempts);
    }

    [Fact]
    public async Task ExecuteAsync_DoesNotRetryTransientFailureForNonIdempotentOperation()
    {
        var attempts = 0;
        var pipeline = CreatePipeline();

        var result = await pipeline.ExecuteAsync(
            OperationType.CreateMailbox,
            _ =>
            {
                attempts++;
                return Task.FromResult(Result<string>.Failure(
                    NormalizedError.Create(ErrorCode.Timeout, "timeout", isTransient: true)));
            },
            CancellationToken.None);

        Assert.True(result.IsFailure);
        Assert.Equal(1, attempts);
    }

    [Fact]
    public async Task ExecuteAsync_OpensCircuitAfterRepeatedTransientFailures()
    {
        var timeProvider = new FakeTimeProvider();
        var pipeline = CreatePipeline(timeProvider);

        for (var i = 0; i < 3; i++)
        {
            var failure = await pipeline.ExecuteAsync(
                OperationType.GetMailboxes,
                _ => Task.FromResult(Result<string>.Failure(
                    NormalizedError.Create(ErrorCode.ServiceUnavailable, "service down", isTransient: true))),
                CancellationToken.None);

            Assert.True(failure.IsFailure);
        }

        var blocked = await pipeline.ExecuteAsync(
            OperationType.GetMailboxes,
            _ => Task.FromResult(Result<string>.Success("should-not-run")),
            CancellationToken.None);

        Assert.True(blocked.IsFailure);
        Assert.Equal(ErrorCode.ServiceUnavailable, blocked.Error!.Code);
        Assert.Contains("Circuit breaker is open", blocked.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ExecuteAsync_AllowsRecoveryAfterCircuitOpenDurationElapses()
    {
        var timeProvider = new FakeTimeProvider();
        var pipeline = CreatePipeline(timeProvider);

        for (var i = 0; i < 3; i++)
        {
            await pipeline.ExecuteAsync(
                OperationType.GetMailboxes,
                _ => Task.FromResult(Result<string>.Failure(
                    NormalizedError.Create(ErrorCode.ServiceUnavailable, "service down", isTransient: true))),
                CancellationToken.None);
        }

        timeProvider.Advance(TimeSpan.FromSeconds(16));

        var recovered = await pipeline.ExecuteAsync(
            OperationType.GetMailboxes,
            _ => Task.FromResult(Result<string>.Success("recovered")),
            CancellationToken.None);

        Assert.True(recovered.IsSuccess);
        Assert.Equal("recovered", recovered.Value);
    }

    [Fact]
    public void IsResilienceEnabled_TreatsMailboxProvisioningCandidatesAsIdempotent()
    {
        var pipeline = CreatePipeline();

        Assert.True(pipeline.IsResilienceEnabled(OperationType.GetMailboxProvisioningCandidates));
    }

    private static WorkerOperationResiliencePipeline CreatePipeline(FakeTimeProvider? timeProvider = null)
    {
        timeProvider ??= new FakeTimeProvider();

        return new WorkerOperationResiliencePipeline(
            retryPolicyFactory: () => new RetryPolicy(new RetryPolicyOptions
            {
                MaxRetries = 3,
                BaseDelayMs = 1,
                MaxDelayMs = 1,
                UseDecorrelatedJitter = false,
                MaxJitter = 0
            }, timeProvider, randomSeed: 123),
            circuitBreakerFactory: () => new CircuitBreaker(new CircuitBreakerOptions
            {
                FailureThreshold = 3,
                OpenDuration = TimeSpan.FromSeconds(15),
                SuccessThresholdInHalfOpen = 1
            }, timeProvider));
    }

    private sealed class FakeTimeProvider : ITimeProvider
    {
        private DateTime _utcNow = new(2026, 3, 11, 12, 0, 0, DateTimeKind.Utc);

        public DateTime UtcNow => _utcNow;

        public void Advance(TimeSpan delta)
        {
            _utcNow = _utcNow.Add(delta);
        }
    }
}

