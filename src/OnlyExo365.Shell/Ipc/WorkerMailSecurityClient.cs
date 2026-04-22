using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;

namespace OnlyExo365.Shell.Ipc;

internal sealed class WorkerMailSecurityClient
{
    private readonly WorkerClientRuntime _runtime;

    public WorkerMailSecurityClient(WorkerClientRuntime runtime)
    {
        _runtime = runtime;
    }

    public Task<Result<GetMailSecurityBaselineResponse>> GetMailSecurityBaselineAsync(
        GetMailSecurityBaselineRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetMailSecurityBaselineResponse>(
            OperationType.GetMailSecurityBaseline,
            request ?? new GetMailSecurityBaselineRequest(),
            eventHandler,
            cancellationToken);

    public Task<Result> UpdateDkimSigningConfigAsync(
        UpdateDkimSigningConfigRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.UpdateDkimSigningConfig, request, eventHandler, cancellationToken);

    public Task<Result> UpdateHostedContentFilterPolicyAsync(
        UpdateHostedContentFilterPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.UpdateHostedContentFilterPolicy, request, eventHandler, cancellationToken);

    public Task<Result> UpdateAntiPhishPolicyAsync(
        UpdateAntiPhishPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.UpdateAntiPhishPolicy, request, eventHandler, cancellationToken);

    public Task<Result> UpdateMalwareFilterPolicyAsync(
        UpdateMalwareFilterPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.UpdateMalwareFilterPolicy, request, eventHandler, cancellationToken);

    public Task<Result> UpdateQuarantinePolicyAsync(
        UpdateQuarantinePolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.UpdateQuarantinePolicy, request, eventHandler, cancellationToken);

    public Task<Result> UpdateHostedOutboundSpamFilterPolicyAsync(
        UpdateHostedOutboundSpamFilterPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.UpdateHostedOutboundSpamFilterPolicy, request, eventHandler, cancellationToken);
}

