using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;

namespace ExchangeAdmin.Tests;

public abstract class TestComplianceWorkerServiceBase : IComplianceWorkerService
{
    public virtual Task<Result<GetComplianceWorkspaceResponse>> GetComplianceWorkspaceAsync(GetComplianceWorkspaceRequest? request = null, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<SearchUnifiedAuditLogResponse>> SearchUnifiedAuditLogAsync(SearchUnifiedAuditLogRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> CreateComplianceSearchAsync(CreateComplianceSearchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> StartComplianceSearchAsync(StartComplianceSearchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> RemoveComplianceSearchAsync(RemoveComplianceSearchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<InvokeComplianceActionResponse>> InvokeComplianceActionAsync(InvokeComplianceActionRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
}
