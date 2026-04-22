using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;

namespace OnlyExo365.Tests;

public abstract class TestComplianceWorkerServiceBase : IComplianceWorkerService
{
    public virtual Task<Result<GetComplianceWorkspaceResponse>> GetComplianceWorkspaceAsync(GetComplianceWorkspaceRequest? request = null, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<SearchUnifiedAuditLogResponse>> SearchUnifiedAuditLogAsync(SearchUnifiedAuditLogRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> CreateComplianceSearchAsync(CreateComplianceSearchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> StartComplianceSearchAsync(StartComplianceSearchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> RemoveComplianceSearchAsync(RemoveComplianceSearchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<InvokeComplianceActionResponse>> InvokeComplianceActionAsync(InvokeComplianceActionRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
}

