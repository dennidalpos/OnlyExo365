using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;

namespace OnlyExo365.Shell.Ipc;

internal sealed class WorkerMailboxClient
{
    private readonly WorkerClientRuntime _runtime;

    public WorkerMailboxClient(WorkerClientRuntime runtime)
    {
        _runtime = runtime;
    }

    public Task<Result<GetMailboxProvisioningCandidatesResponse>> GetMailboxProvisioningCandidatesAsync(
        GetMailboxProvisioningCandidatesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetMailboxProvisioningCandidatesResponse>(
            OperationType.GetMailboxProvisioningCandidates,
            request,
            eventHandler,
            cancellationToken);

    public Task<Result<GetMailboxesResponse>> GetMailboxesAsync(
        GetMailboxesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetMailboxesResponse>(OperationType.GetMailboxes, request, eventHandler, cancellationToken);

    public Task<Result<GetDeletedMailboxesResponse>> GetDeletedMailboxesAsync(
        GetDeletedMailboxesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetDeletedMailboxesResponse>(OperationType.GetDeletedMailboxes, request, eventHandler, cancellationToken);

    public Task<Result<MailboxDetailsDto>> GetMailboxDetailsAsync(
        GetMailboxDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<MailboxDetailsDto>(OperationType.GetMailboxDetails, request, eventHandler, cancellationToken);

    public Task<Result<GetRetentionPoliciesResponse>> GetRetentionPoliciesAsync(
        GetRetentionPoliciesRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetRetentionPoliciesResponse>(
            OperationType.GetRetentionPolicies,
            request ?? new GetRetentionPoliciesRequest(),
            eventHandler,
            cancellationToken);

    public Task<Result> SetRetentionPolicyAsync(
        SetRetentionPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.SetRetentionPolicy, request, eventHandler, cancellationToken);

    public Task<Result<MailboxPermissionsDto>> GetMailboxPermissionsAsync(
        string identity,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<MailboxPermissionsDto>(
            OperationType.GetMailboxPermissions,
            new GetMailboxPermissionsRequest { Identity = identity },
            eventHandler,
            cancellationToken);

    public Task<Result<GetMailboxFolderPermissionsResponse>> GetMailboxFolderPermissionsAsync(
        GetMailboxFolderPermissionsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetMailboxFolderPermissionsResponse>(
            OperationType.GetMailboxFolderPermissions,
            request,
            eventHandler,
            cancellationToken);

    public Task<Result> SetMailboxPermissionAsync(
        SetMailboxPermissionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.SetMailboxPermission, request, eventHandler, cancellationToken);

    public Task<Result> SetMailboxFolderPermissionAsync(
        SetMailboxFolderPermissionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.SetMailboxFolderPermission, request, eventHandler, cancellationToken);

    public Task<Result<ApplyPermissionsDeltaPlanResponse>> ApplyPermissionsDeltaPlanAsync(
        ApplyPermissionsDeltaPlanRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<ApplyPermissionsDeltaPlanResponse>(
            OperationType.ApplyPermissionsDeltaPlan,
            request,
            eventHandler,
            cancellationToken);

    public Task<Result> UpdateMailboxSettingsAsync(
        UpdateMailboxSettingsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.UpdateMailboxSettings, request, eventHandler, cancellationToken);

    public Task<Result> SetMailboxAutoReplyConfigurationAsync(
        SetMailboxAutoReplyConfigurationRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.SetMailboxAutoReplyConfiguration, request, eventHandler, cancellationToken);

    public Task<Result> ConvertMailboxToSharedAsync(
        ConvertMailboxToSharedRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.ConvertMailboxToShared, request, eventHandler, cancellationToken);

    public Task<Result> ConvertMailboxToRegularAsync(
        ConvertMailboxToRegularRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.ConvertMailboxToRegular, request, eventHandler, cancellationToken);

    public Task<Result<RestoreMailboxResponse>> RestoreMailboxAsync(
        RestoreMailboxRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<RestoreMailboxResponse>(OperationType.RestoreMailbox, request, eventHandler, cancellationToken);

    public Task<Result<GetMailboxSpaceReportResponse>> GetMailboxSpaceReportAsync(
        GetMailboxSpaceReportRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetMailboxSpaceReportResponse>(OperationType.GetMailboxSpaceReport, request, eventHandler, cancellationToken);

    public Task<Result<GetMailboxAccessReportResponse>> GetMailboxAccessReportAsync(
        GetMailboxAccessReportRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetMailboxAccessReportResponse>(OperationType.GetMailboxAccessReport, request, eventHandler, cancellationToken);

    public Task<Result> CreateMailboxAsync(
        CreateMailboxRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.CreateMailbox, request, eventHandler, cancellationToken);
}

