using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;

namespace ExchangeAdmin.Tests;

public abstract class TestDistributionListsWorkerServiceBase : IDistributionListsWorkerService
{
    public virtual Task<Result<GetDistributionListsResponse>> GetDistributionListsAsync(GetDistributionListsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<DistributionListDetailsDto>> GetDistributionListDetailsAsync(GetDistributionListDetailsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GroupMembersPageDto>> GetGroupMembersAsync(GetGroupMembersRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> ModifyGroupMemberAsync(ModifyGroupMemberRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> SetDistributionListSettingsAsync(SetDistributionListSettingsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> CreateDistributionListAsync(CreateDistributionListRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<PreviewDynamicGroupMembersResponse>> PreviewDynamicGroupMembersAsync(PreviewDynamicGroupMembersRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetAcceptedDomainsResponse>> GetAcceptedDomainsAsync(GetAcceptedDomainsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
}
