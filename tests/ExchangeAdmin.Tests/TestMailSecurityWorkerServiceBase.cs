using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;

namespace ExchangeAdmin.Tests;

public abstract class TestMailSecurityWorkerServiceBase : IMailSecurityWorkerService
{
    public virtual Task<Result<GetMailSecurityBaselineResponse>> GetMailSecurityBaselineAsync(GetMailSecurityBaselineRequest? request = null, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpdateDkimSigningConfigAsync(UpdateDkimSigningConfigRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpdateHostedContentFilterPolicyAsync(UpdateHostedContentFilterPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpdateAntiPhishPolicyAsync(UpdateAntiPhishPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpdateMalwareFilterPolicyAsync(UpdateMalwareFilterPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpdateQuarantinePolicyAsync(UpdateQuarantinePolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpdateHostedOutboundSpamFilterPolicyAsync(UpdateHostedOutboundSpamFilterPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
}
