using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;

namespace ExchangeAdmin.Infrastructure.Ipc;

internal sealed class WorkerGroupClient
{
    private readonly WorkerClientRuntime _runtime;

    public WorkerGroupClient(WorkerClientRuntime runtime)
    {
        _runtime = runtime;
    }

    public Task<Result<GetDistributionListsResponse>> GetDistributionListsAsync(
        GetDistributionListsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetDistributionListsResponse>(OperationType.GetDistributionLists, request, eventHandler, cancellationToken);

    public Task<Result<DistributionListDetailsDto>> GetDistributionListDetailsAsync(
        GetDistributionListDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<DistributionListDetailsDto>(OperationType.GetDistributionListDetails, request, eventHandler, cancellationToken);

    public Task<Result<GroupMembersPageDto>> GetGroupMembersAsync(
        GetGroupMembersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GroupMembersPageDto>(OperationType.GetGroupMembers, request, eventHandler, cancellationToken);

    public Task<Result> ModifyGroupMemberAsync(
        ModifyGroupMemberRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.ModifyGroupMember, request, eventHandler, cancellationToken);

    public Task<Result> SetDistributionListSettingsAsync(
        SetDistributionListSettingsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.SetDistributionListSettings, request, eventHandler, cancellationToken);

    public Task<Result> CreateDistributionListAsync(
        CreateDistributionListRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.CreateDistributionList, request, eventHandler, cancellationToken);

    public Task<Result<PreviewDynamicGroupMembersResponse>> PreviewDynamicGroupMembersAsync(
        PreviewDynamicGroupMembersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<PreviewDynamicGroupMembersResponse>(
            OperationType.PreviewDynamicGroupMembers,
            request,
            eventHandler,
            cancellationToken);
}
