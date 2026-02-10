using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;
using ExchangeAdmin.Infrastructure.Ipc;

namespace ExchangeAdmin.Application.Services;

             
                                            
              
public interface IWorkerService
{
                 
                                        
                  
    WorkerConnectionState ConnectionState { get; }

                 
                               
                  
    WorkerStatus Status { get; }

                 
                                                         
                  
    CapabilityMapDto? Capabilities { get; }

                 
                                       
                  
    event EventHandler<WorkerConnectionState>? StateChanged;

                 
                                   
                  
    event EventHandler<EventEnvelope>? EventReceived;

                 
                                   
                  
    event EventHandler<CapabilityMapDto>? CapabilitiesUpdated;

    #region Worker Lifecycle

                 
                                 
                  
    Task<bool> StartWorkerAsync(CancellationToken cancellationToken = default);

                 
                                
                  
    Task StopWorkerAsync();

                 
                                   
                  
    Task<bool> RestartWorkerAsync(CancellationToken cancellationToken = default);

                 
                                      
                  
    void KillWorker();

    #endregion

    #region Connection

                 
                                                 
                  
    Task<Result<ConnectionStatusDto>> ConnectExchangeAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

                 
                                        
                  
    Task<Result> DisconnectExchangeAsync(CancellationToken cancellationToken = default);

                 
                                               
                  
    Task<Result<ConnectionStatusDto>> GetConnectionStatusAsync(CancellationToken cancellationToken = default);

                 
                                      
                  
    Task<Result<CapabilityMapDto>> DetectCapabilitiesAsync(
        bool forceRefresh = false,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    #endregion

    #region Dashboard

                 
                                 
                  
    Task<Result<DashboardStatsDto>> GetDashboardStatsAsync(
        GetDashboardStatsRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    #endregion

    #region Mailboxes

                 
                                     
                  
    Task<Result<GetMailboxesResponse>> GetMailboxesAsync(
        GetMailboxesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetDeletedMailboxesResponse>> GetDeletedMailboxesAsync(
        GetDeletedMailboxesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

                 
                            
                  
    Task<Result<MailboxDetailsDto>> GetMailboxDetailsAsync(
        GetMailboxDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

                 
                               
                  
    Task<Result<GetRetentionPoliciesResponse>> GetRetentionPoliciesAsync(
        GetRetentionPoliciesRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

                 
                                     
                  
    Task<Result> SetRetentionPolicyAsync(
        SetRetentionPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

                 
                                
                  
    Task<Result<MailboxPermissionsDto>> GetMailboxPermissionsAsync(
        string identity,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

                 
                                      
                  
    Task<Result> SetMailboxPermissionAsync(
        SetMailboxPermissionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

                 
                                     
                  
    Task<Result<ApplyPermissionsDeltaPlanResponse>> ApplyPermissionsDeltaPlanAsync(
        ApplyPermissionsDeltaPlanRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

                 
                            
                  
    Task<Result> SetMailboxFeatureAsync(
        SetMailboxFeatureRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpdateMailboxSettingsAsync(
        UpdateMailboxSettingsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> SetMailboxAutoReplyConfigurationAsync(
        SetMailboxAutoReplyConfigurationRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> ConvertMailboxToSharedAsync(
        ConvertMailboxToSharedRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> ConvertMailboxToRegularAsync(
        ConvertMailboxToRegularRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<RestoreMailboxResponse>> RestoreMailboxAsync(
        RestoreMailboxRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetMailboxSpaceReportResponse>> GetMailboxSpaceReportAsync(
        GetMailboxSpaceReportRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    #endregion

    #region Distribution Lists

                 
                                           
                  
    Task<Result<GetDistributionListsResponse>> GetDistributionListsAsync(
        GetDistributionListsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

                 
                                      
                  
    Task<Result<DistributionListDetailsDto>> GetDistributionListDetailsAsync(
        GetDistributionListDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

                 
                                      
                  
    Task<Result<GroupMembersPageDto>> GetGroupMembersAsync(
        GetGroupMembersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

                 
                                   
                  
    Task<Result> ModifyGroupMemberAsync(
        ModifyGroupMemberRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> SetDistributionListSettingsAsync(
        SetDistributionListSettingsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

                 
                                                   
                  
    Task<Result<PreviewDynamicGroupMembersResponse>> PreviewDynamicGroupMembersAsync(
        PreviewDynamicGroupMembersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    #endregion

    #region Message Trace

    Task<Result<GetMessageTraceResponse>> GetMessageTraceAsync(
        GetMessageTraceRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetMessageTraceDetailsResponse>> GetMessageTraceDetailsAsync(
        GetMessageTraceDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    #endregion


    #region Mail Flow

    Task<Result<GetTransportRulesResponse>> GetTransportRulesAsync(
        GetTransportRulesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> SetTransportRuleStateAsync(
        SetTransportRuleStateRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpsertTransportRuleAsync(
        UpsertTransportRuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> RemoveTransportRuleAsync(
        RemoveTransportRuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<TestTransportRuleResponse>> TestTransportRuleAsync(
        TestTransportRuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetConnectorsResponse>> GetConnectorsAsync(
        GetConnectorsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetAcceptedDomainsResponse>> GetAcceptedDomainsAsync(
        GetAcceptedDomainsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpsertConnectorAsync(
        UpsertConnectorRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> RemoveConnectorAsync(
        RemoveConnectorRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpsertAcceptedDomainAsync(
        UpsertAcceptedDomainRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> RemoveAcceptedDomainAsync(
        RemoveAcceptedDomainRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    #endregion

    #region Licenses

    Task<Result<GetUserLicensesResponse>> GetUserLicensesAsync(
        GetUserLicensesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> SetUserLicenseAsync(
        SetUserLicenseRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetAvailableLicensesResponse>> GetAvailableLicensesAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    #endregion

    #region System

    Task<Result<PrerequisiteStatusDto>> CheckPrerequisitesAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<InstallModuleResponse>> InstallModuleAsync(
        InstallModuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    #endregion

    #region Demo




    Task<Result<DemoOperationResponse>> RunDemoOperationAsync(
        DemoOperationRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    #endregion




    Task CancelOperationAsync(string correlationId);
}
