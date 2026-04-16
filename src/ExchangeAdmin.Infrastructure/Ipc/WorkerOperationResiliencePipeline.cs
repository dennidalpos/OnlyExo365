using System.Collections.Concurrent;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Errors;
using ExchangeAdmin.Domain.Resilience;
using ExchangeAdmin.Domain.Results;

namespace ExchangeAdmin.Infrastructure.Ipc;

internal sealed class WorkerOperationResiliencePipeline
{
    private static readonly HashSet<OperationType> IdempotentOperations =
    [
        OperationType.GetConnectionStatus,
        OperationType.DetectCapabilities,
        OperationType.GetDashboardStats,
        OperationType.GetContacts,
        OperationType.GetContactDetails,
        OperationType.GetResourceMailboxes,
        OperationType.GetResourceMailboxDetails,
        OperationType.GetPublicFolders,
        OperationType.GetPublicFolderDetails,
        OperationType.GetMobileDevices,
        OperationType.GetMobileDeviceMailboxPolicies,
        OperationType.GetMigrationBatches,
        OperationType.GetMigrationBatchDetails,
        OperationType.GetRoleGroups,
        OperationType.GetRoleGroupDetails,
        OperationType.GetMailboxProvisioningCandidates,
        OperationType.GetMailboxes,
        OperationType.GetDeletedMailboxes,
        OperationType.GetMailboxDetails,
        OperationType.GetRetentionPolicies,
        OperationType.GetMailboxPermissions,
        OperationType.GetMailboxFolderPermissions,
        OperationType.GetMailboxSpaceReport,
        OperationType.GetMailboxAccessReport,
        OperationType.GetDistributionLists,
        OperationType.GetDistributionListDetails,
        OperationType.GetGroupMembers,
        OperationType.PreviewDynamicGroupMembers,
        OperationType.GetMessageTrace,
        OperationType.GetMessageTraceDetails,
        OperationType.GetTransportRules,
        OperationType.GetConnectors,
        OperationType.GetAcceptedDomains,
        OperationType.GetRemoteDomains,
        OperationType.GetOrganizationRelationships,
        OperationType.GetAddressLists,
        OperationType.GetAddressBookPolicies,
        OperationType.GetOfflineAddressBooks,
        OperationType.GetSharingPolicies,
        OperationType.GetUserLicenses,
        OperationType.GetAvailableLicenses,
        OperationType.CheckPrerequisites
    ];

    private readonly ConcurrentDictionary<OperationType, CircuitBreaker> _breakers = new();
    private readonly Func<RetryPolicy> _retryPolicyFactory;
    private readonly Func<CircuitBreaker> _circuitBreakerFactory;

    public WorkerOperationResiliencePipeline(
        Func<RetryPolicy>? retryPolicyFactory = null,
        Func<CircuitBreaker>? circuitBreakerFactory = null)
    {
        _retryPolicyFactory = retryPolicyFactory ?? (() => new RetryPolicy(new RetryPolicyOptions
        {
            MaxRetries = 3,
            BaseDelayMs = 250,
            MaxDelayMs = 2000,
            BackoffFactor = 2.0,
            UseDecorrelatedJitter = true,
            MaxJitter = 0.2
        }));
        _circuitBreakerFactory = circuitBreakerFactory ?? (() => new CircuitBreaker(new CircuitBreakerOptions
        {
            FailureThreshold = 3,
            OpenDuration = TimeSpan.FromSeconds(15),
            SuccessThresholdInHalfOpen = 1
        }));
    }

    public bool IsResilienceEnabled(OperationType operation) => IdempotentOperations.Contains(operation);

    public Task<Result<T>> ExecuteAsync<T>(
        OperationType operation,
        Func<CancellationToken, Task<Result<T>>> action,
        CancellationToken cancellationToken)
    {
        if (!IsResilienceEnabled(operation))
        {
            return action(cancellationToken);
        }

        return ExecuteWithResilienceAsync(operation, action, cancellationToken);
    }

    public Task<Result> ExecuteAsync(
        OperationType operation,
        Func<CancellationToken, Task<Result>> action,
        CancellationToken cancellationToken)
    {
        if (!IsResilienceEnabled(operation))
        {
            return action(cancellationToken);
        }

        return ExecuteWithResilienceAsync(operation, action, cancellationToken);
    }

    private async Task<Result<T>> ExecuteWithResilienceAsync<T>(
        OperationType operation,
        Func<CancellationToken, Task<Result<T>>> action,
        CancellationToken cancellationToken)
    {
        var breaker = GetBreaker(operation);
        if (!breaker.CanExecute())
        {
            return CreateCircuitOpenFailure<T>(operation, breaker.RemainingOpenTime);
        }

        var result = await _retryPolicyFactory().ExecuteResultAsync(action, cancellationToken).ConfigureAwait(false);
        ApplyBreakerOutcome(breaker, result);
        return result;
    }

    private async Task<Result> ExecuteWithResilienceAsync(
        OperationType operation,
        Func<CancellationToken, Task<Result>> action,
        CancellationToken cancellationToken)
    {
        var breaker = GetBreaker(operation);
        if (!breaker.CanExecute())
        {
            return CreateCircuitOpenFailure(operation, breaker.RemainingOpenTime);
        }

        var result = await _retryPolicyFactory().ExecuteResultAsync(action, cancellationToken).ConfigureAwait(false);
        ApplyBreakerOutcome(breaker, result);
        return result;
    }

    private CircuitBreaker GetBreaker(OperationType operation)
        => _breakers.GetOrAdd(operation, _ => _circuitBreakerFactory());

    private static void ApplyBreakerOutcome<T>(CircuitBreaker breaker, Result<T> result)
    {
        if (result.IsSuccess)
        {
            breaker.RecordSuccess();
            return;
        }

        if (result.WasCancelled || result.Error == null)
        {
            return;
        }

        if (ShouldCountAsAvailabilityFailure(result.Error))
        {
            breaker.RecordFailure();
        }
    }

    private static void ApplyBreakerOutcome(CircuitBreaker breaker, Result result)
    {
        if (result.IsSuccess)
        {
            breaker.RecordSuccess();
            return;
        }

        if (result.WasCancelled || result.Error == null)
        {
            return;
        }

        if (ShouldCountAsAvailabilityFailure(result.Error))
        {
            breaker.RecordFailure();
        }
    }

    private static bool ShouldCountAsAvailabilityFailure(NormalizedError error)
        => error.IsTransient && !NonRetryableErrors.IsNonRetryable(error.Code);

    private static Result<T> CreateCircuitOpenFailure<T>(OperationType operation, TimeSpan remainingOpenTime)
        => Result<T>.Failure(NormalizedError.Create(
            ErrorCode.ServiceUnavailable,
            $"Circuit breaker is open for {operation}. Retry in {Math.Ceiling(Math.Max(remainingOpenTime.TotalSeconds, 0))} seconds.",
            isTransient: true,
            retryAfter: remainingOpenTime));

    private static Result CreateCircuitOpenFailure(OperationType operation, TimeSpan remainingOpenTime)
        => Result.Failure(NormalizedError.Create(
            ErrorCode.ServiceUnavailable,
            $"Circuit breaker is open for {operation}. Retry in {Math.Ceiling(Math.Max(remainingOpenTime.TotalSeconds, 0))} seconds.",
            isTransient: true,
            retryAfter: remainingOpenTime));
}
