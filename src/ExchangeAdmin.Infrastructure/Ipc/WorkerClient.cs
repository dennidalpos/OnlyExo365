using System.Text.Json;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Errors;
using ExchangeAdmin.Domain.Results;

namespace ExchangeAdmin.Infrastructure.Ipc;

             
                                            
                                                          
              
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

                 
                                                 
                  
    public async Task<Result<ConnectionStatusDto>> ConnectExchangeAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        var result = await ExecuteOperationAsync<ConnectionStatusDto>(
            OperationType.ConnectExchangeInteractive,
            null,
            eventHandler,
            cancellationToken);

                                                               
        if (result.IsSuccess && result.Value?.State == ConnectionState.Connected)
        {
            _ = DetectCapabilitiesAsync(forceRefresh: true, cancellationToken: CancellationToken.None);
        }

        return result;
    }

                 
                                        
                  
    public async Task<Result> DisconnectExchangeAsync(CancellationToken cancellationToken = default)
    {
        var response = await SendRequestInternalAsync(
            OperationType.DisconnectExchange,
            null,
            null,
            cancellationToken);

                                           
        _capabilities = null;

        if (response.WasCancelled)
            return Result.Cancelled();

        if (!response.Success)
            return Result.Failure(NormalizedError.FromDto(response.Error!));

        return Result.Success();
    }

                 
                                      
                  
    public async Task<Result<ConnectionStatusDto>> GetConnectionStatusAsync(CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<ConnectionStatusDto>(
            OperationType.GetConnectionStatus,
            null,
            null,
            cancellationToken);
    }

                 
                                      
                  
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

    public async Task<Result<GetDeletedMailboxesResponse>> GetDeletedMailboxesAsync(
        GetDeletedMailboxesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<GetDeletedMailboxesResponse>(
            OperationType.GetDeletedMailboxes,
            request,
            eventHandler,
            cancellationToken);
    }

                 
                            
                  
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

                 
                               
                  
    public async Task<Result<GetRetentionPoliciesResponse>> GetRetentionPoliciesAsync(
        GetRetentionPoliciesRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<GetRetentionPoliciesResponse>(
            OperationType.GetRetentionPolicies,
            request ?? new GetRetentionPoliciesRequest(),
            eventHandler,
            cancellationToken);
    }

                 
                                     
                  
    public async Task<Result> SetRetentionPolicyAsync(
        SetRetentionPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        var response = await SendRequestInternalAsync(
            OperationType.SetRetentionPolicy,
            request,
            eventHandler,
            cancellationToken);

        if (response.WasCancelled)
            return Result.Cancelled();

        if (!response.Success)
            return Result.Failure(NormalizedError.FromDto(response.Error!));

        return Result.Success();
    }

                 
                                
                  
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

                 
                                         
                  
    public async Task<Result<RestoreMailboxResponse>> RestoreMailboxAsync(
        RestoreMailboxRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<RestoreMailboxResponse>(
            OperationType.RestoreMailbox,
            request,
            eventHandler,
            cancellationToken);
    }

                 
                                
                 
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

    #region Message Trace

    public async Task<Result<GetMessageTraceResponse>> GetMessageTraceAsync(
        GetMessageTraceRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<GetMessageTraceResponse>(
            OperationType.GetMessageTrace,
            request,
            eventHandler,
            cancellationToken);
    }

    #endregion

    #region Licenses

    public async Task<Result<GetUserLicensesResponse>> GetUserLicensesAsync(
        GetUserLicensesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<GetUserLicensesResponse>(
            OperationType.GetUserLicenses,
            request,
            eventHandler,
            cancellationToken);
    }

    public async Task<Result> SetUserLicenseAsync(
        SetUserLicenseRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        var response = await SendRequestInternalAsync(
            OperationType.SetUserLicense,
            request,
            eventHandler,
            cancellationToken);

        if (response.WasCancelled)
            return Result.Cancelled();

        if (!response.Success)
            return Result.Failure(NormalizedError.FromDto(response.Error!));

        return Result.Success();
    }

    public async Task<Result<GetAvailableLicensesResponse>> GetAvailableLicensesAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<GetAvailableLicensesResponse>(
            OperationType.GetAvailableLicenses,
            null,
            eventHandler,
            cancellationToken);
    }

    #endregion

    #region System

    public async Task<Result<PrerequisiteStatusDto>> CheckPrerequisitesAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<PrerequisiteStatusDto>(
            OperationType.CheckPrerequisites,
            null,
            eventHandler,
            cancellationToken);
    }

    public async Task<Result<InstallModuleResponse>> InstallModuleAsync(
        InstallModuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteOperationAsync<InstallModuleResponse>(
            OperationType.InstallModule,
            request,
            eventHandler,
            cancellationToken);
    }

    #endregion

    #region Demo


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

                             
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var correlationId = request.CorrelationId;

                                    
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
