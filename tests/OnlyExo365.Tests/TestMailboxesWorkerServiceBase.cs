using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;

namespace OnlyExo365.Tests;

public abstract class TestMailboxesWorkerServiceBase : IMailboxesWorkerService
{
    public virtual Task<Result<GetMailboxProvisioningCandidatesResponse>> GetMailboxProvisioningCandidatesAsync(GetMailboxProvisioningCandidatesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetMailboxesResponse>> GetMailboxesAsync(GetMailboxesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetDeletedMailboxesResponse>> GetDeletedMailboxesAsync(GetDeletedMailboxesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<MailboxDetailsDto>> GetMailboxDetailsAsync(GetMailboxDetailsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetRetentionPoliciesResponse>> GetRetentionPoliciesAsync(GetRetentionPoliciesRequest? request = null, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> SetRetentionPolicyAsync(SetRetentionPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<ApplyPermissionsDeltaPlanResponse>> ApplyPermissionsDeltaPlanAsync(ApplyPermissionsDeltaPlanRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetMailboxFolderPermissionsResponse>> GetMailboxFolderPermissionsAsync(GetMailboxFolderPermissionsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> SetMailboxFolderPermissionAsync(SetMailboxFolderPermissionRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpdateMailboxSettingsAsync(UpdateMailboxSettingsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> SetMailboxAutoReplyConfigurationAsync(SetMailboxAutoReplyConfigurationRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> ConvertMailboxToSharedAsync(ConvertMailboxToSharedRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> ConvertMailboxToRegularAsync(ConvertMailboxToRegularRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<RestoreMailboxResponse>> RestoreMailboxAsync(RestoreMailboxRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetMailboxSpaceReportResponse>> GetMailboxSpaceReportAsync(GetMailboxSpaceReportRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetMailboxAccessReportResponse>> GetMailboxAccessReportAsync(GetMailboxAccessReportRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> CreateMailboxAsync(CreateMailboxRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetAcceptedDomainsResponse>> GetAcceptedDomainsAsync(GetAcceptedDomainsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetUserLicensesResponse>> GetUserLicensesAsync(GetUserLicensesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> SetUserLicenseAsync(SetUserLicenseRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetUsageLocationSuggestionResponse>> GetUsageLocationSuggestionAsync(GetUsageLocationSuggestionRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> SetUserUsageLocationAsync(SetUserUsageLocationRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetAvailableLicensesResponse>> GetAvailableLicensesAsync(Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
}

