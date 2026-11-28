using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;

namespace OnlyExo365.Shell.Ipc;

internal sealed class WorkerComplianceClient
{
    private readonly WorkerClientRuntime _runtime;

    public WorkerComplianceClient(WorkerClientRuntime runtime)
    {
        _runtime = runtime;
    }

    public Task<Result<GetComplianceWorkspaceResponse>> GetComplianceWorkspaceAsync(
        GetComplianceWorkspaceRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetComplianceWorkspaceResponse>(
            OperationType.GetComplianceWorkspace,
            request ?? new GetComplianceWorkspaceRequest(),
            eventHandler,
            cancellationToken);

    public Task<Result<SearchUnifiedAuditLogResponse>> SearchUnifiedAuditLogAsync(
        SearchUnifiedAuditLogRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<SearchUnifiedAuditLogResponse>(
            OperationType.SearchUnifiedAuditLog,
            request,
            eventHandler,
            cancellationToken);

    public Task<Result> CreateComplianceSearchAsync(
        CreateComplianceSearchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.CreateComplianceSearch, request, eventHandler, cancellationToken);

    public Task<Result> StartComplianceSearchAsync(
        StartComplianceSearchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.StartComplianceSearch, request, eventHandler, cancellationToken);

    public Task<Result> RemoveComplianceSearchAsync(
        RemoveComplianceSearchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.RemoveComplianceSearch, request, eventHandler, cancellationToken);

    public Task<Result<InvokeComplianceActionResponse>> InvokeComplianceActionAsync(
        InvokeComplianceActionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<InvokeComplianceActionResponse>(
            OperationType.InvokeComplianceAction,
            request,
            eventHandler,
            cancellationToken);
}

