using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

internal sealed class ExoMailboxCommands
{
    private readonly ExoMailboxListingCommands _listingCommands;
    private readonly ExoMailboxDetailCommands _detailCommands;
    private readonly ExoMailboxPolicyCommands _policyCommands;
    private readonly ExoMailboxLifecycleCommands _lifecycleCommands;
    private readonly ExoMailboxLicenseCommands _mailboxLicenseCommands;

    public ExoMailboxCommands(
        PowerShellEngine engine,
        CapabilityDetector capabilityDetector,
        ExoMailboxReportingCommands mailboxReportingCommands,
        ExoMailboxLicenseCommands mailboxLicenseCommands)
    {
        _listingCommands = new ExoMailboxListingCommands(engine);
        _detailCommands = new ExoMailboxDetailCommands(engine, capabilityDetector, mailboxReportingCommands, mailboxLicenseCommands);
        _policyCommands = new ExoMailboxPolicyCommands(engine, capabilityDetector);
        _lifecycleCommands = new ExoMailboxLifecycleCommands(engine);
        _mailboxLicenseCommands = mailboxLicenseCommands;
    }

    public Task<GetMailboxProvisioningCandidatesResponse> GetMailboxProvisioningCandidatesAsync(
        GetMailboxProvisioningCandidatesRequest request,
        CancellationToken cancellationToken = default)
        => _mailboxLicenseCommands.GetMailboxProvisioningCandidatesAsync(request, cancellationToken);

    public Task<GetMailboxesResponse> GetMailboxesAsync(
        GetMailboxesRequest request,
        Action<string, string>? onLog = null,
        Action<MailboxListItemDto>? onPartialOutput = null,
        CancellationToken cancellationToken = default)
        => _listingCommands.GetMailboxesAsync(request, onLog, onPartialOutput, cancellationToken);

    internal static string BuildGetMailboxesScript(
        int skip,
        int pageSize,
        string filterParam,
        string escapedSearch,
        string sortProperty,
        string sortDirection,
        bool useWindowedLoad)
        => ExoMailboxScriptFactory.BuildGetMailboxesScript(
            skip,
            pageSize,
            filterParam,
            escapedSearch,
            sortProperty,
            sortDirection,
            useWindowedLoad);

    internal static int CalculatePageWindowSize(int skip, int pageSize)
        => ExoMailboxScriptFactory.CalculatePageWindowSize(skip, pageSize);

    public Task<GetDeletedMailboxesResponse> GetDeletedMailboxesAsync(
        GetDeletedMailboxesRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _listingCommands.GetDeletedMailboxesAsync(request, onLog, cancellationToken);

    public Task<MailboxDetailsDto> GetMailboxDetailsAsync(
        GetMailboxDetailsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _detailCommands.GetMailboxDetailsAsync(request, onLog, cancellationToken);

    public Task<List<RetentionPolicySummaryDto>> GetRetentionPoliciesAsync(
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _policyCommands.GetRetentionPoliciesAsync(onLog, cancellationToken);

    public Task SetMailboxSettingsAsync(
        UpdateMailboxSettingsRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => _policyCommands.SetMailboxSettingsAsync(request, onLog, cancellationToken);

    internal static string BuildSetMailboxSettingsScript(UpdateMailboxSettingsRequest request)
        => ExoMailboxScriptFactory.BuildSetMailboxSettingsScript(request);

    public Task SetRetentionPolicyAsync(
        SetRetentionPolicyRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => _policyCommands.SetRetentionPolicyAsync(request, onLog, cancellationToken);

    public Task SetMailboxAutoReplyConfigurationAsync(
        SetMailboxAutoReplyConfigurationRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => _policyCommands.SetMailboxAutoReplyConfigurationAsync(request, onLog, cancellationToken);

    public Task CreateMailboxAsync(
        CreateMailboxRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => _lifecycleCommands.CreateMailboxAsync(request, onLog, cancellationToken);

    internal static (string Script, Dictionary<string, object>? Parameters) BuildCreateMailboxCommand(CreateMailboxRequest request)
        => ExoMailboxScriptFactory.BuildCreateMailboxCommand(request);

    public Task ConvertMailboxToSharedAsync(
        ConvertMailboxToSharedRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => _lifecycleCommands.ConvertMailboxToSharedAsync(request, onLog, cancellationToken);

    public Task ConvertMailboxToRegularAsync(
        ConvertMailboxToRegularRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => _lifecycleCommands.ConvertMailboxToRegularAsync(request, onLog, cancellationToken);

    public Task<RestoreMailboxResponse> RestoreMailboxAsync(
        RestoreMailboxRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => _lifecycleCommands.RestoreMailboxAsync(request, onLog, cancellationToken);

    internal static (string Script, Dictionary<string, object> Parameters) BuildRestoreMailboxCommand(RestoreMailboxRequest request)
        => ExoMailboxScriptFactory.BuildRestoreMailboxCommand(request);
}

