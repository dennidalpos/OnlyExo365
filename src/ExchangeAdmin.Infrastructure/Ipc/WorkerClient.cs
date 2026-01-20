using System.Text.Json;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Errors;
using ExchangeAdmin.Domain.Results;

namespace ExchangeAdmin.Infrastructure.Ipc;

/// <summary>
/// High-level client for worker operations.
/// Wraps IpcClient with strong typing and error handling.
/// </summary>
public class WorkerClient : IAsyncDisposable
{
    private readonly WorkerSupervisor _supervisor;
    private CapabilityMapDto? _capabilities;

    public event EventHandler<WorkerConnectionState>? StateChanged;
    public event EventHandler<EventEnvelope>? EventReceived;
    public event EventHandler<CapabilityMapDto>? CapabilitiesUpdated;

    public WorkerConnectionState State => _supervisor.State;
    public WorkerStatus Status => _supervisor.GetStatus();
    public CapabilityMapDto? Capabilities => _capabilities;

    public WorkerClient(WorkerSupervisorOptions? options = null)
    {
        _supervisor = new WorkerSupervisor(options);
        _supervisor.StateChanged += (s, e) => StateChanged?.Invoke(this, e);
        _supervisor.EventReceived += (s, e) => EventReceived?.Invoke(this, e);
    }

    public Task<bool> StartWorkerAsync(CancellationToken cancellationToken = default)
        => _supervisor.StartAsync(cancellationToken);

    public Task StopWorkerAsync()
        => _supervisor.StopAsync();

    public Task<bool> RestartWorkerAsync(CancellationToken cancellationToken = default)
        => _supervisor.RestartAsync(cancellationToken);

    public void KillWorker()
        => _supervisor.KillWorker();

    #region Connection

    /// <summary>
    /// Connect to Exchange Online (interactive).
    /// </summary>
    public async Task<Result<ConnectionStatusDto>> ConnectExchangeAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        var result = await ExecuteOperationAsync<ConnectionStatusDto>(
            OperationType.ConnectExchangeInteractive,
            null,
            eventHandler,
            cancellationToken);

        // Auto-detect capabilities after successful connection
        if (result.IsSuccess && result.Value?.State == ConnectionState.Connected)
        {
            _ = DetectCapabilitiesAsync(forceRefresh: true, cancellationToken: CancellationToken.None);
        }

        return result;
    }

    /// <summary>
    /// Disconnect from Exchange Online.
    /// </summary>
    public async Task<Result> DisconnectExchangeAsync(CancellationToken cancellationToken = default)
    {
        var response = await SendRequestInternalAsync(
            OperationType.DisconnectExchange,
            null,
            null,
            cancellationToken);

        // Clear capabilities on disconnect
        _capabilities = null;

        if (response.WasCancelled)
            return Result.Cancelled();

        if (!response.Success)
            return Result.Failure(NormalizedError.FromDto(response.Error!));

        return Result.Success();
    }

    /// <summary>
    /// Get current connection status.
    /// </summary>
    public async Task<Result<ConnectionStatusDto>> GetConnectionStatusAsync(CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<ConnectionStatusDto>(
            OperationType.GetConnectionStatus,
            null,
            null,
            cancellationToken);
    }

    /// <summary>
    /// Detect available capabilities.
    /// </summary>
    public async Task<Result<CapabilityMapDto>> DetectCapabilitiesAsync(
        bool forceRefresh = false,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        var result = await ExecuteOperationAsync<CapabilityMapDto>(
            OperationType.DetectCapabilities,
            new DetectCapabilitiesRequest { ForceRefresh = forceRefresh },
            eventHandler,
            cancellationToken);

        if (result.IsSuccess && result.Value != null)
        {
            _capabilities = result.Value;
            CapabilitiesUpdated?.Invoke(this, result.Value);
        }

        return result;
    }

    #endregion

    #region Dashboard

    /// <summary>
    /// Get dashboard statistics.
    /// </summary>
    public async Task<Result<DashboardStatsDto>> GetDashboardStatsAsync(
        GetDashboardStatsRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<DashboardStatsDto>(
            OperationType.GetDashboardStats,
            request ?? new GetDashboardStatsRequest(),
            eventHandler,
            cancellationToken);
    }

    #endregion

    #region Mailboxes

    /// <summary>
    /// Get mailbox list with paging.
    /// </summary>
    public async Task<Result<GetMailboxesResponse>> GetMailboxesAsync(
        GetMailboxesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<GetMailboxesResponse>(
            OperationType.GetMailboxes,
            request,
            eventHandler,
            cancellationToken);
    }

    /// <summary>
    /// Get mailbox details.
    /// </summary>
    public async Task<Result<MailboxDetailsDto>> GetMailboxDetailsAsync(
        GetMailboxDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<MailboxDetailsDto>(
            OperationType.GetMailboxDetails,
            request,
            eventHandler,
            cancellationToken);
    }

    /// <summary>
    /// Get mailbox permissions.
    /// </summary>
    public async Task<Result<MailboxPermissionsDto>> GetMailboxPermissionsAsync(
        string identity,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<MailboxPermissionsDto>(
            OperationType.GetMailboxPermissions,
            new GetMailboxPermissionsRequest { Identity = identity },
            eventHandler,
            cancellationToken);
    }

    /// <summary>
    /// Set mailbox permission.
    /// </summary>
    public async Task<Result> SetMailboxPermissionAsync(
        SetMailboxPermissionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        var response = await SendRequestInternalAsync(
            OperationType.SetMailboxPermission,
            request,
            eventHandler,
            cancellationToken);

        if (response.WasCancelled)
            return Result.Cancelled();

        if (!response.Success)
            return Result.Failure(NormalizedError.FromDto(response.Error!));

        return Result.Success();
    }

    /// <summary>
    /// Apply permissions delta plan.
    /// </summary>
    public async Task<Result<ApplyPermissionsDeltaPlanResponse>> ApplyPermissionsDeltaPlanAsync(
        ApplyPermissionsDeltaPlanRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<ApplyPermissionsDeltaPlanResponse>(
            OperationType.ApplyPermissionsDeltaPlan,
            request,
            eventHandler,
            cancellationToken);
    }

    /// <summary>
    /// Set mailbox feature.
    /// </summary>
    public async Task<Result> SetMailboxFeatureAsync(
        SetMailboxFeatureRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        var response = await SendRequestInternalAsync(
            OperationType.SetMailboxFeature,
            request,
            eventHandler,
            cancellationToken);

        if (response.WasCancelled)
            return Result.Cancelled();

        if (!response.Success)
            return Result.Failure(NormalizedError.FromDto(response.Error!));

        return Result.Success();
    }

    /// <summary>
    /// Update mailbox settings.
    /// </summary>
    public async Task<Result> UpdateMailboxSettingsAsync(
        UpdateMailboxSettingsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        var response = await SendRequestInternalAsync(
            OperationType.UpdateMailboxSettings,
            request,
            eventHandler,
            cancellationToken);

        if (response.WasCancelled)
            return Result.Cancelled();

        if (!response.Success)
            return Result.Failure(NormalizedError.FromDto(response.Error!));

        return Result.Success();
    }

    /// <summary>
    /// Update mailbox auto-reply configuration.
    /// </summary>
    public async Task<Result> SetMailboxAutoReplyConfigurationAsync(
        SetMailboxAutoReplyConfigurationRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        var response = await SendRequestInternalAsync(
            OperationType.SetMailboxAutoReplyConfiguration,
            request,
            eventHandler,
            cancellationToken);

        if (response.WasCancelled)
            return Result.Cancelled();

        if (!response.Success)
            return Result.Failure(NormalizedError.FromDto(response.Error!));

        return Result.Success();
    }

    /// <summary>
    /// Convert mailbox to shared mailbox.
    /// </summary>
    public async Task<Result> ConvertMailboxToSharedAsync(
        ConvertMailboxToSharedRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        var response = await SendRequestInternalAsync(
            OperationType.ConvertMailboxToShared,
            request,
            eventHandler,
            cancellationToken);

        if (response.WasCancelled)
            return Result.Cancelled();

        if (!response.Success)
            return Result.Failure(NormalizedError.FromDto(response.Error!));

        return Result.Success();
    }

    /// <summary>
    /// Convert mailbox to regular mailbox.
    /// </summary>
    public async Task<Result> ConvertMailboxToRegularAsync(
        ConvertMailboxToRegularRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        var response = await SendRequestInternalAsync(
            OperationType.ConvertMailboxToRegular,
            request,
            eventHandler,
            cancellationToken);

        if (response.WasCancelled)
            return Result.Cancelled();

        if (!response.Success)
            return Result.Failure(NormalizedError.FromDto(response.Error!));

        return Result.Success();
    }

    /// <summary>
    /// Get mailbox space report.
    /// </summary>
    public async Task<Result<GetMailboxSpaceReportResponse>> GetMailboxSpaceReportAsync(
        GetMailboxSpaceReportRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<GetMailboxSpaceReportResponse>(
            OperationType.GetMailboxSpaceReport,
            request,
            eventHandler,
            cancellationToken);
    }

    #endregion

    #region Distribution Lists

    /// <summary>
    /// Get distribution lists with paging.
    /// </summary>
    public async Task<Result<GetDistributionListsResponse>> GetDistributionListsAsync(
        GetDistributionListsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<GetDistributionListsResponse>(
            OperationType.GetDistributionLists,
            request,
            eventHandler,
            cancellationToken);
    }

    /// <summary>
    /// Get distribution list details.
    /// </summary>
    public async Task<Result<DistributionListDetailsDto>> GetDistributionListDetailsAsync(
        GetDistributionListDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<DistributionListDetailsDto>(
            OperationType.GetDistributionListDetails,
            request,
            eventHandler,
            cancellationToken);
    }

    /// <summary>
    /// Get group members with paging.
    /// </summary>
    public async Task<Result<GroupMembersPageDto>> GetGroupMembersAsync(
        GetGroupMembersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<GroupMembersPageDto>(
            OperationType.GetGroupMembers,
            request,
            eventHandler,
            cancellationToken);
    }

    /// <summary>
    /// Modify group member.
    /// </summary>
    public async Task<Result> ModifyGroupMemberAsync(
        ModifyGroupMemberRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        var response = await SendRequestInternalAsync(
            OperationType.ModifyGroupMember,
            request,
            eventHandler,
            cancellationToken);

        if (response.WasCancelled)
            return Result.Cancelled();

        if (!response.Success)
            return Result.Failure(NormalizedError.FromDto(response.Error!));

        return Result.Success();
    }

    /// <summary>
    /// Update distribution list settings.
    /// </summary>
    public async Task<Result> SetDistributionListSettingsAsync(
        SetDistributionListSettingsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        var response = await SendRequestInternalAsync(
            OperationType.SetDistributionListSettings,
            request,
            eventHandler,
            cancellationToken);

        if (response.WasCancelled)
            return Result.Cancelled();

        if (!response.Success)
            return Result.Failure(NormalizedError.FromDto(response.Error!));

        return Result.Success();
    }

    /// <summary>
    /// Preview dynamic distribution group members.
    /// </summary>
    public async Task<Result<PreviewDynamicGroupMembersResponse>> PreviewDynamicGroupMembersAsync(
        PreviewDynamicGroupMembersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<PreviewDynamicGroupMembersResponse>(
            OperationType.PreviewDynamicGroupMembers,
            request,
            eventHandler,
            cancellationToken);
    }

    #endregion

    #region Demo

    /// <summary>
    /// Run demo long operation with streaming.
    /// </summary>
    public async Task<Result<DemoOperationResponse>> RunDemoOperationAsync(
        DemoOperationRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<DemoOperationResponse>(
            OperationType.DemoLongOperation,
            request,
            eventHandler,
            cancellationToken);
    }

    #endregion

    /// <summary>
    /// Cancel an operation in progress.
    /// </summary>
    public async Task CancelOperationAsync(string correlationId)
    {
        if (_supervisor.State != WorkerConnectionState.Connected)
            return;

        await _supervisor.IpcClient.SendCancelAsync(correlationId);
    }

    private async Task<Result<TResponse>> ExecuteOperationAsync<TResponse>(
        OperationType operation,
        object? payload,
        Action<EventEnvelope>? eventHandler,
        CancellationToken cancellationToken)
    {
        var response = await SendRequestInternalAsync(operation, payload, eventHandler, cancellationToken);

        if (response.WasCancelled)
            return Result<TResponse>.Cancelled();

        if (!response.Success)
            return Result<TResponse>.Failure(NormalizedError.FromDto(response.Error!));

        if (response.Payload == null)
            return Result<TResponse>.Success(default!);

        var result = JsonMessageSerializer.ExtractPayload<TResponse>(response.Payload);
        return Result<TResponse>.Success(result!);
    }

    private async Task<ResponseEnvelope> SendRequestInternalAsync(
        OperationType operation,
        object? payload,
        Action<EventEnvelope>? eventHandler,
        CancellationToken cancellationToken)
    {
        if (_supervisor.State != WorkerConnectionState.Connected)
        {
            return new ResponseEnvelope
            {
                Success = false,
                Error = new NormalizedErrorDto
                {
                    Code = ErrorCode.WorkerNotRunning,
                    Message = "Worker is not running",
                    IsTransient = false
                }
            };
        }

        var request = new RequestEnvelope
        {
            Operation = operation,
            Payload = payload != null ? JsonMessageSerializer.ToJsonElement(payload) : null
        };

        // Setup cancellation
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var correlationId = request.CorrelationId;

        // Register for cancellation
        var registration = cancellationToken.Register(async () =>
        {
            await _supervisor.IpcClient.SendCancelAsync(correlationId);
        });

        try
        {
            return await _supervisor.IpcClient.SendRequestAsync(request, eventHandler, linkedCts.Token);
        }
        catch (OperationCanceledException)
        {
            return new ResponseEnvelope
            {
                CorrelationId = correlationId,
                Success = false,
                WasCancelled = true
            };
        }
        finally
        {
            await registration.DisposeAsync();
        }
    }

    public async ValueTask DisposeAsync()
    {
        await _supervisor.DisposeAsync();
    }
}
