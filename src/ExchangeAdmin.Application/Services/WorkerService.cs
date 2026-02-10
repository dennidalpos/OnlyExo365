using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;
using ExchangeAdmin.Infrastructure.Ipc;

namespace ExchangeAdmin.Application.Services;

             
                                  
              
public class WorkerService : IWorkerService, IAsyncDisposable
{
    private readonly WorkerClient _client;

    public WorkerConnectionState ConnectionState => _client.State;
    public WorkerStatus Status => _client.Status;
    public CapabilityMapDto? Capabilities => _client.Capabilities;

    public event EventHandler<WorkerConnectionState>? StateChanged;
    public event EventHandler<EventEnvelope>? EventReceived;
    public event EventHandler<CapabilityMapDto>? CapabilitiesUpdated;

    public WorkerService(WorkerSupervisorOptions? options = null)
    {
        _client = new WorkerClient(options);
        _client.StateChanged += (s, e) => StateChanged?.Invoke(this, e);
        _client.EventReceived += (s, e) => EventReceived?.Invoke(this, e);
        _client.CapabilitiesUpdated += (s, e) => CapabilitiesUpdated?.Invoke(this, e);
    }

    #region Worker Lifecycle

    public Task<bool> StartWorkerAsync(CancellationToken cancellationToken = default)
        => _client.StartWorkerAsync(cancellationToken);

    public Task StopWorkerAsync()
        => _client.StopWorkerAsync();

    public Task<bool> RestartWorkerAsync(CancellationToken cancellationToken = default)
        => _client.RestartWorkerAsync(cancellationToken);

    public void KillWorker()
        => _client.KillWorker();

    #endregion

    #region Connection

    public Task<Result<ConnectionStatusDto>> ConnectExchangeAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.ConnectExchangeAsync(eventHandler, cancellationToken);

    public Task<Result> DisconnectExchangeAsync(CancellationToken cancellationToken = default)
        => _client.DisconnectExchangeAsync(cancellationToken);

    public Task<Result<ConnectionStatusDto>> GetConnectionStatusAsync(CancellationToken cancellationToken = default)
        => _client.GetConnectionStatusAsync(cancellationToken);

    public Task<Result<CapabilityMapDto>> DetectCapabilitiesAsync(
        bool forceRefresh = false,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.DetectCapabilitiesAsync(forceRefresh, eventHandler, cancellationToken);

    #endregion

    #region Dashboard

    public Task<Result<DashboardStatsDto>> GetDashboardStatsAsync(
        GetDashboardStatsRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetDashboardStatsAsync(request, eventHandler, cancellationToken);

    #endregion

    #region Mailboxes

    public Task<Result<GetMailboxesResponse>> GetMailboxesAsync(
        GetMailboxesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMailboxesAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetDeletedMailboxesResponse>> GetDeletedMailboxesAsync(
        GetDeletedMailboxesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetDeletedMailboxesAsync(request, eventHandler, cancellationToken);

    public Task<Result<MailboxDetailsDto>> GetMailboxDetailsAsync(
        GetMailboxDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMailboxDetailsAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetRetentionPoliciesResponse>> GetRetentionPoliciesAsync(
        GetRetentionPoliciesRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetRetentionPoliciesAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetRetentionPolicyAsync(
        SetRetentionPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SetRetentionPolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result<MailboxPermissionsDto>> GetMailboxPermissionsAsync(
        string identity,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMailboxPermissionsAsync(identity, eventHandler, cancellationToken);

    public Task<Result> SetMailboxPermissionAsync(
        SetMailboxPermissionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SetMailboxPermissionAsync(request, eventHandler, cancellationToken);

    public Task<Result<ApplyPermissionsDeltaPlanResponse>> ApplyPermissionsDeltaPlanAsync(
        ApplyPermissionsDeltaPlanRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.ApplyPermissionsDeltaPlanAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetMailboxFeatureAsync(
        SetMailboxFeatureRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SetMailboxFeatureAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpdateMailboxSettingsAsync(
        UpdateMailboxSettingsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpdateMailboxSettingsAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetMailboxAutoReplyConfigurationAsync(
        SetMailboxAutoReplyConfigurationRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SetMailboxAutoReplyConfigurationAsync(request, eventHandler, cancellationToken);

    public Task<Result> ConvertMailboxToSharedAsync(
        ConvertMailboxToSharedRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.ConvertMailboxToSharedAsync(request, eventHandler, cancellationToken);

    public Task<Result> ConvertMailboxToRegularAsync(
        ConvertMailboxToRegularRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.ConvertMailboxToRegularAsync(request, eventHandler, cancellationToken);

    public Task<Result<RestoreMailboxResponse>> RestoreMailboxAsync(
        RestoreMailboxRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.RestoreMailboxAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMailboxSpaceReportResponse>> GetMailboxSpaceReportAsync(
        GetMailboxSpaceReportRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMailboxSpaceReportAsync(request, eventHandler, cancellationToken);

    #endregion

    #region Distribution Lists

    public Task<Result<GetDistributionListsResponse>> GetDistributionListsAsync(
        GetDistributionListsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetDistributionListsAsync(request, eventHandler, cancellationToken);

    public Task<Result<DistributionListDetailsDto>> GetDistributionListDetailsAsync(
        GetDistributionListDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetDistributionListDetailsAsync(request, eventHandler, cancellationToken);

    public Task<Result<GroupMembersPageDto>> GetGroupMembersAsync(
        GetGroupMembersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetGroupMembersAsync(request, eventHandler, cancellationToken);

    public Task<Result> ModifyGroupMemberAsync(
        ModifyGroupMemberRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.ModifyGroupMemberAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetDistributionListSettingsAsync(
        SetDistributionListSettingsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SetDistributionListSettingsAsync(request, eventHandler, cancellationToken);

    public Task<Result<PreviewDynamicGroupMembersResponse>> PreviewDynamicGroupMembersAsync(
        PreviewDynamicGroupMembersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.PreviewDynamicGroupMembersAsync(request, eventHandler, cancellationToken);

    #endregion

    #region Message Trace

    public Task<Result<GetMessageTraceResponse>> GetMessageTraceAsync(
        GetMessageTraceRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMessageTraceAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMessageTraceDetailsResponse>> GetMessageTraceDetailsAsync(
        GetMessageTraceDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMessageTraceDetailsAsync(request, eventHandler, cancellationToken);

    #endregion

    #region Mail Flow

    public Task<Result<GetTransportRulesResponse>> GetTransportRulesAsync(
        GetTransportRulesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetTransportRulesAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetTransportRuleStateAsync(
        SetTransportRuleStateRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SetTransportRuleStateAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetConnectorsResponse>> GetConnectorsAsync(
        GetConnectorsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetConnectorsAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetAcceptedDomainsResponse>> GetAcceptedDomainsAsync(
        GetAcceptedDomainsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetAcceptedDomainsAsync(request, eventHandler, cancellationToken);

    #endregion

    #region Licenses

    public Task<Result<GetUserLicensesResponse>> GetUserLicensesAsync(
        GetUserLicensesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetUserLicensesAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetUserLicenseAsync(
        SetUserLicenseRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SetUserLicenseAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetAvailableLicensesResponse>> GetAvailableLicensesAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetAvailableLicensesAsync(eventHandler, cancellationToken);

    #endregion

    #region System

    public Task<Result<PrerequisiteStatusDto>> CheckPrerequisitesAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.CheckPrerequisitesAsync(eventHandler, cancellationToken);

    public Task<Result<InstallModuleResponse>> InstallModuleAsync(
        InstallModuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.InstallModuleAsync(request, eventHandler, cancellationToken);

    #endregion

    #region Demo

    public Task<Result<DemoOperationResponse>> RunDemoOperationAsync(
        DemoOperationRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.RunDemoOperationAsync(request, eventHandler, cancellationToken);

    #endregion

    public Task CancelOperationAsync(string correlationId)
        => _client.CancelOperationAsync(correlationId);

    public async ValueTask DisposeAsync()
    {
        await _client.DisposeAsync();
    }
}
