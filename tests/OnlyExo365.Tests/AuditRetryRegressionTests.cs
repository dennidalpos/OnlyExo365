using OnlyExo365.Contracts.Errors;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;
using OnlyExo365.Shell.Resilience;

namespace OnlyExo365.Tests;

public sealed class AuditRetryRegressionTests
{
    [Fact]
    public async Task RetryAfterAboveLocalCap_WaitsUntilCancellation()
    {
        var attempts = 0;
        var policy = new RetryPolicy(new RetryPolicyOptions { MaxRetries = 2, BaseDelayMs = 1, MaxDelayMs = 1 });
        using var cancellation = new CancellationTokenSource(TimeSpan.FromMilliseconds(100));
        var result = await policy.ExecuteResultAsync(_ =>
        {
            attempts++;
            return Task.FromResult(attempts == 1
                ? Result<string>.Failure(NormalizedError.Create(ErrorCode.ServiceUnavailable, "throttled", isTransient: true, retryAfter: TimeSpan.FromSeconds(60)))
                : Result<string>.Success("retried too early"));
        }, cancellation.Token);
        Assert.True(result.WasCancelled);
        Assert.Equal(1, attempts);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CancellationDuringExceptionBackoff_ReturnsCancelledResult(bool structuredResult)
    {
        var policy = new RetryPolicy(new RetryPolicyOptions { MaxRetries = 2, BaseDelayMs = 60000, MaxDelayMs = 60000 });
        using var cancellation = new CancellationTokenSource(TimeSpan.FromMilliseconds(100));
        if (structuredResult)
        {
            var result = await policy.ExecuteResultAsync<string>(_ => throw new TimeoutException("timeout"), cancellation.Token);
            Assert.True(result.WasCancelled);
        }
        else
        {
            var result = await policy.ExecuteAsync<string>(_ => throw new TimeoutException("timeout"), cancellation.Token);
            Assert.True(result.WasCancelled);
            Assert.Equal(1, result.Attempts);
        }
    }
}
