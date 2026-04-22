using System.Text;
using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

internal static class ExoMailboxScriptFactory
{
    internal static string BuildGetMailboxesScript(
        int skip,
        int pageSize,
        string filterParam,
        string escapedSearch,
        string sortProperty,
        string sortDirection,
        bool useWindowedLoad)
    {
        if (useWindowedLoad)
        {
            var pageWindowSize = CalculatePageWindowSize(skip, pageSize);
            return $@"
$pageWindowSize = {pageWindowSize}
$allMailboxes = @(Get-Mailbox -ResultSize $pageWindowSize {filterParam})
$hasMore = @($allMailboxes).Count -gt {pageSize}
$pagedMailboxes = $allMailboxes | Select-Object -Skip {skip} -First {pageSize}

@{{
    TotalCount = {skip} + @($pagedMailboxes).Count + $(if ($hasMore) {{ 1 }} else {{ 0 }})
    HasMore = $hasMore
    IsTotalCountExact = $false
    Mailboxes = @($pagedMailboxes | ForEach-Object {{
        @{{
            Identity = $_.Identity.ToString()
            Guid = $_.ExchangeGuid.ToString()
            DisplayName = $_.DisplayName
            PrimarySmtpAddress = $_.PrimarySmtpAddress.ToString()
            UserPrincipalName = if ($_.UserPrincipalName) {{ $_.UserPrincipalName.ToString() }} else {{ '' }}
            RecipientType = $_.RecipientType.ToString()
            RecipientTypeDetails = $_.RecipientTypeDetails.ToString()
            Alias = $_.Alias
            IsInactiveMailbox = $_.IsInactiveMailbox
        }}
    }})
}}";
        }

        var searchFilter = string.IsNullOrWhiteSpace(escapedSearch)
            ? string.Empty
            : $@"
$allMailboxes = $allMailboxes | Where-Object {{
    $_.DisplayName -like '*{escapedSearch}*' -or
    $_.PrimarySmtpAddress -like '*{escapedSearch}*' -or
    $_.Alias -like '*{escapedSearch}*'
}}";

        return $@"
$allMailboxes = Get-Mailbox -ResultSize Unlimited {filterParam}{searchFilter}
$allMailboxes = $allMailboxes | Sort-Object {sortProperty} {sortDirection}
$totalCount = @($allMailboxes).Count
$pagedMailboxes = $allMailboxes | Select-Object -Skip {skip} -First {pageSize}

@{{
    TotalCount = $totalCount
    HasMore = ({skip} + @($pagedMailboxes).Count) -lt $totalCount
    IsTotalCountExact = $true
    Mailboxes = @($pagedMailboxes | ForEach-Object {{
        @{{
            Identity = $_.Identity.ToString()
            Guid = $_.ExchangeGuid.ToString()
            DisplayName = $_.DisplayName
            PrimarySmtpAddress = $_.PrimarySmtpAddress.ToString()
            UserPrincipalName = if ($_.UserPrincipalName) {{ $_.UserPrincipalName.ToString() }} else {{ '' }}
            RecipientType = $_.RecipientType.ToString()
            RecipientTypeDetails = $_.RecipientTypeDetails.ToString()
            Alias = $_.Alias
            IsInactiveMailbox = $_.IsInactiveMailbox
        }}
    }})
}}";
    }

    internal static int CalculatePageWindowSize(int skip, int pageSize)
        => Math.Max(0, skip) + Math.Max(1, pageSize) + 1;

    internal static string BuildGetDeletedMailboxesScript(GetDeletedMailboxesRequest request)
    {
        var includeSoftDeleted = request.IncludeSoftDeleted ? "$true" : "$false";
        var includeInactive = request.IncludeInactive ? "$true" : "$false";
        var escapedSearch = EscapePs(request.SearchQuery);

        return $@"
$softDeleted = @()
$inactive = @()

if ({includeSoftDeleted}) {{
    $softDeleted = Get-Mailbox -SoftDeletedMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue
}}

if ({includeInactive}) {{
    $inactive = Get-Mailbox -InactiveMailboxOnly -ResultSize Unlimited -ErrorAction SilentlyContinue
}}

$allMailboxes = @()
$softDeleted | ForEach-Object {{
    $allMailboxes += [pscustomobject]@{{
        Identity = $_.Identity.ToString()
        DisplayName = $_.DisplayName
        UserPrincipalName = $_.UserPrincipalName
        PrimarySmtpAddress = if ($_.PrimarySmtpAddress) {{ $_.PrimarySmtpAddress.ToString() }} else {{ '' }}
        RecipientTypeDetails = $_.RecipientTypeDetails.ToString()
        Alias = $_.Alias
        DeletionType = 'SoftDeleted'
    }}
}}

$inactive | ForEach-Object {{
    $allMailboxes += [pscustomobject]@{{
        Identity = $_.Identity.ToString()
        DisplayName = $_.DisplayName
        UserPrincipalName = $_.UserPrincipalName
        PrimarySmtpAddress = if ($_.PrimarySmtpAddress) {{ $_.PrimarySmtpAddress.ToString() }} else {{ '' }}
        RecipientTypeDetails = $_.RecipientTypeDetails.ToString()
        Alias = $_.Alias
        DeletionType = 'Inactive'
    }}
}}

$allMailboxes = $allMailboxes | Group-Object Identity | ForEach-Object {{
    $_.Group |
        Sort-Object @{{
            Expression = {{
                switch ($_.DeletionType) {{
                    'Inactive' {{ 0 }}
                    'SoftDeleted' {{ 1 }}
                    'HardDeleted' {{ 2 }}
                    default {{ 99 }}
                }}
            }}
        }} |
        Select-Object -First 1
}}

$searchQuery = '{escapedSearch}'
if (-not [string]::IsNullOrWhiteSpace($searchQuery)) {{
    $allMailboxes = $allMailboxes | Where-Object {{
        $_.DisplayName -like ""*$searchQuery*"" -or
        $_.PrimarySmtpAddress -like ""*$searchQuery*"" -or
        $_.UserPrincipalName -like ""*$searchQuery*"" -or
        $_.Alias -like ""*$searchQuery*""
    }}
}}

$allMailboxes = $allMailboxes | Sort-Object DisplayName
$totalCount = @($allMailboxes).Count
$pagedMailboxes = $allMailboxes | Select-Object -Skip {request.Skip} -First {request.PageSize}

@{{
    TotalCount = $totalCount
    Mailboxes = @($pagedMailboxes | ForEach-Object {{
        @{{
            Identity = $_.Identity
            DisplayName = $_.DisplayName
            UserPrincipalName = $_.UserPrincipalName
            PrimarySmtpAddress = $_.PrimarySmtpAddress
            RecipientTypeDetails = $_.RecipientTypeDetails
            Alias = $_.Alias
            DeletionType = $_.DeletionType
        }}
    }})
}}";
    }

    internal static string BuildGetMailboxDetailsScript(string identity)
    {
        var escapedIdentity = EscapePs(identity);

        return $@"
$ErrorActionPreference = 'Stop'

function Get-BytesFromSize($size) {{
    if ($null -eq $size) {{ return $null }}
    $text = $size.ToString()
    if ($text -match '\(([^)]+)\s+byte[s]?\)') {{
        $numeric = ($Matches[1] -replace '[^\d]', '')
        if (-not [string]::IsNullOrWhiteSpace($numeric)) {{
            return [long]$numeric
        }}
    }}
    return $null
}}

$mbx = Get-Mailbox -Identity '{escapedIdentity}' -ErrorAction Stop
if ($null -eq $mbx) {{
    throw ""Mailbox not found: {escapedIdentity}""
}}

@{{
    Identity = $mbx.Identity.ToString()
    Guid = $mbx.ExchangeGuid.ToString()
    DisplayName = $mbx.DisplayName
    PrimarySmtpAddress = $mbx.PrimarySmtpAddress.ToString()
    UserPrincipalName = $mbx.UserPrincipalName
    Alias = $mbx.Alias
    RecipientType = $mbx.RecipientType.ToString()
    RecipientTypeDetails = $mbx.RecipientTypeDetails.ToString()
    EmailAddresses = @($mbx.EmailAddresses | ForEach-Object {{ $_.ToString() }})
    RetentionPolicy = $mbx.RetentionPolicy
    WhenCreated = $mbx.WhenCreated
    WhenMailboxCreated = $mbx.WhenMailboxCreated
    ArchiveEnabled = if ($mbx.ArchiveDatabase) {{ $true }} else {{ $false }}
    ArchiveName = $mbx.ArchiveName
    ArchiveGuid = $mbx.ArchiveGuid.ToString()
    ArchiveStatus = $mbx.ArchiveStatus.ToString()
    LitigationHoldEnabled = $mbx.LitigationHoldEnabled
    LitigationHoldDate = $mbx.LitigationHoldDate
    LitigationHoldOwner = $mbx.LitigationHoldOwner
    LitigationHoldDuration = if ($mbx.LitigationHoldDuration) {{ $mbx.LitigationHoldDuration.ToString() }} else {{ $null }}
    AuditEnabled = $mbx.AuditEnabled
    AuditLogAgeLimit = if ($mbx.AuditLogAgeLimit) {{ $mbx.AuditLogAgeLimit.ToString() }} else {{ $null }}
    AuditAdmin = @($mbx.AuditAdmin)
    AuditDelegate = @($mbx.AuditDelegate)
    AuditOwner = @($mbx.AuditOwner)
    ForwardingAddress = if ($mbx.ForwardingAddress) {{ $mbx.ForwardingAddress.ToString() }} else {{ $null }}
    ForwardingSmtpAddress = if ($mbx.ForwardingSmtpAddress) {{ $mbx.ForwardingSmtpAddress.ToString() }} else {{ $null }}
    DeliverToMailboxAndForward = $mbx.DeliverToMailboxAndForward
    ProhibitSendQuota = if ($mbx.ProhibitSendQuota) {{ $mbx.ProhibitSendQuota.ToString() }} else {{ $null }}
    ProhibitSendQuotaBytes = Get-BytesFromSize $mbx.ProhibitSendQuota
    ProhibitSendReceiveQuota = if ($mbx.ProhibitSendReceiveQuota) {{ $mbx.ProhibitSendReceiveQuota.ToString() }} else {{ $null }}
    ProhibitSendReceiveQuotaBytes = Get-BytesFromSize $mbx.ProhibitSendReceiveQuota
    IssueWarningQuota = if ($mbx.IssueWarningQuota) {{ $mbx.IssueWarningQuota.ToString() }} else {{ $null }}
    IssueWarningQuotaBytes = Get-BytesFromSize $mbx.IssueWarningQuota
    MaxSendSize = if ($mbx.MaxSendSize) {{ $mbx.MaxSendSize.ToString() }} else {{ $null }}
    MaxReceiveSize = if ($mbx.MaxReceiveSize) {{ $mbx.MaxReceiveSize.ToString() }} else {{ $null }}
    RetentionHoldEnabled = $mbx.RetentionHoldEnabled
    SingleItemRecoveryEnabled = $mbx.SingleItemRecoveryEnabled
    RetainDeletedItemsFor = if ($mbx.RetainDeletedItemsFor) {{ $mbx.RetainDeletedItemsFor.ToString() }} else {{ $null }}
    HiddenFromAddressListsEnabled = $mbx.HiddenFromAddressListsEnabled
}}";
    }

    internal static string BuildGetMailboxCasSettingsScript(string identity)
    {
        var escapedIdentity = EscapePs(identity);

        return $@"
$ErrorActionPreference = 'Stop'

$cas = Get-CASMailbox -Identity '{escapedIdentity}' -ErrorAction Stop
@{{
    OwaEnabled = $cas.OWAEnabled
    ActiveSyncEnabled = $cas.ActiveSyncEnabled
    MapiEnabled = $cas.MAPIEnabled
    PopEnabled = $cas.PopEnabled
    ImapEnabled = $cas.ImapEnabled
    SmtpClientAuthenticationDisabled = $cas.SmtpClientAuthenticationDisabled
}}";
    }

    internal static string BuildGetMailboxStatisticsScript(string identity)
    {
        var escapedIdentity = EscapePs(identity);

        return $@"
$ErrorActionPreference = 'SilentlyContinue'

function Get-BytesFromSize($size) {{
    if ($null -eq $size) {{ return $null }}
    $text = $size.ToString()
    if ($text -match '\(([^)]+)\s+byte[s]?\)') {{
        $numeric = ($Matches[1] -replace '[^\d]', '')
        if (-not [string]::IsNullOrWhiteSpace($numeric)) {{
            return [long]$numeric
        }}
    }}
    return $null
}}

try {{
    $stats = Get-MailboxStatistics -Identity '{escapedIdentity}' -ErrorAction SilentlyContinue
    if ($null -eq $stats) {{
        return $null
    }}
    @{{
        TotalItemSize = $stats.TotalItemSize.ToString()
        TotalItemSizeBytes = Get-BytesFromSize $stats.TotalItemSize
        ItemCount = $stats.ItemCount
        DeletedItemCount = $stats.DeletedItemCount
        TotalDeletedItemSize = if ($stats.TotalDeletedItemSize) {{ $stats.TotalDeletedItemSize.ToString() }} else {{ $null }}
        LastLogonTime = $stats.LastLogonTime
        LastLogoffTime = $stats.LastLogoffTime
    }}
}}
catch {{
    $null
}}";
    }

    internal static string BuildGetInboxRulesScript(string identity)
    {
        var escapedIdentity = EscapePs(identity);

        return $@"
$ErrorActionPreference = 'SilentlyContinue'

try {{
    $rules = Get-InboxRule -Mailbox '{escapedIdentity}' -ErrorAction SilentlyContinue
    if ($null -eq $rules) {{
        @()
        return
    }}
    @($rules | ForEach-Object {{
        @{{
            Name = $_.Name
            RuleIdentity = $_.RuleIdentity.ToString()
            Enabled = $_.Enabled
            Priority = $_.Priority
            Description = $_.Description
            ForwardTo = @($_.ForwardTo | ForEach-Object {{ $_.ToString() }})
            ForwardAsAttachmentTo = @($_.ForwardAsAttachmentTo | ForEach-Object {{ $_.ToString() }})
            RedirectTo = @($_.RedirectTo | ForEach-Object {{ $_.ToString() }})
            DeleteMessage = $_.DeleteMessage
            MoveToFolder = $_.MoveToFolder
        }}
    }})
}}
catch {{
    @()
}}";
    }

    internal static string BuildGetAutoReplyConfigurationScript(string identity)
    {
        var escapedIdentity = EscapePs(identity);

        return $@"
$ErrorActionPreference = 'SilentlyContinue'

try {{
    $config = Get-MailboxAutoReplyConfiguration -Identity '{escapedIdentity}' -ErrorAction SilentlyContinue
    if ($null -eq $config) {{
        return $null
    }}
    @{{
        AutoReplyState = $config.AutoReplyState.ToString()
        StartTime = $config.StartTime
        EndTime = $config.EndTime
        InternalMessage = $config.InternalMessage
        ExternalMessage = $config.ExternalMessage
        ExternalAudience = $config.ExternalAudience.ToString()
    }}
}}
catch {{
    $null
}}";
    }

    internal static string BuildGetRetentionPoliciesScript()
    {
        return @"
$policies = Get-RetentionPolicy
@($policies | ForEach-Object {
    $requiresArchive = $false
    $tagLinks = @($_.RetentionPolicyTagLinks)
    foreach ($tagLink in $tagLinks) {
        try {
            $tag = Get-RetentionPolicyTag -Identity $tagLink -ErrorAction Stop
            if ($tag.RetentionAction -eq 'MoveToArchive') {
                $requiresArchive = $true
                break
            }
        }
        catch {
        }
    }
    @{
        Id = $_.Guid.ToString()
        Name = $_.Name
        Description = $_.Description
        RequiresArchive = $requiresArchive
    }
})";
    }

    internal static string BuildSetMailboxSettingsScript(UpdateMailboxSettingsRequest request, FeatureCapabilitiesDto? capabilities = null)
    {
        var escapedIdentity = EscapePs(request.Identity);
        var scriptBuilder = new StringBuilder();

        if (request.ArchiveEnabled.HasValue)
        {
            if (request.ArchiveEnabled.Value)
            {
                scriptBuilder.AppendLine($"Enable-Mailbox -Identity '{escapedIdentity}' -Archive");
            }
            else
            {
                scriptBuilder.AppendLine($"Disable-Mailbox -Identity '{escapedIdentity}' -Archive -Confirm:$false");
            }
        }

        var setMailboxParams = new List<string>();

        if (request.LitigationHoldEnabled.HasValue)
        {
            setMailboxParams.Add($"-LitigationHoldEnabled ${request.LitigationHoldEnabled.Value.ToString().ToLowerInvariant()}");
        }

        if (request.AuditEnabled.HasValue)
        {
            setMailboxParams.Add($"-AuditEnabled ${request.AuditEnabled.Value.ToString().ToLowerInvariant()}");
        }

        if (request.SingleItemRecoveryEnabled.HasValue)
        {
            setMailboxParams.Add($"-SingleItemRecoveryEnabled ${request.SingleItemRecoveryEnabled.Value.ToString().ToLowerInvariant()}");
        }

        if (request.RetentionHoldEnabled.HasValue)
        {
            setMailboxParams.Add($"-RetentionHoldEnabled ${request.RetentionHoldEnabled.Value.ToString().ToLowerInvariant()}");
        }

        if (request.ForwardingAddress != null)
        {
            setMailboxParams.Add($"-ForwardingAddress {FormatNullableString(request.ForwardingAddress)}");
        }

        if (request.ForwardingSmtpAddress != null)
        {
            setMailboxParams.Add($"-ForwardingSmtpAddress {FormatNullableString(request.ForwardingSmtpAddress)}");
        }

        if (request.DeliverToMailboxAndForward.HasValue)
        {
            setMailboxParams.Add($"-DeliverToMailboxAndForward ${request.DeliverToMailboxAndForward.Value.ToString().ToLowerInvariant()}");
        }

        if (request.PrimarySmtpAddress != null || request.ProxyAddresses != null)
        {
            var primarySmtpAddress = NormalizePrimarySmtpAddress(request.PrimarySmtpAddress);
            if (string.IsNullOrWhiteSpace(primarySmtpAddress))
            {
                throw new InvalidOperationException("PrimarySmtpAddress is required when updating mailbox proxy addresses.");
            }

            var emailAddresses = BuildMailboxEmailAddresses(primarySmtpAddress, request.ProxyAddresses);
            setMailboxParams.Add($"-EmailAddresses {ExoRequestSanitizer.FormatStringArrayParameter(emailAddresses)}");
        }

        if (request.HiddenFromAddressListsEnabled.HasValue)
        {
            setMailboxParams.Add($"-HiddenFromAddressListsEnabled ${request.HiddenFromAddressListsEnabled.Value.ToString().ToLowerInvariant()}");
        }

        if (request.IssueWarningQuota != null)
        {
            setMailboxParams.Add($"-IssueWarningQuota {FormatNullableString(request.IssueWarningQuota)}");
        }

        if (request.ProhibitSendQuota != null)
        {
            setMailboxParams.Add($"-ProhibitSendQuota {FormatNullableString(request.ProhibitSendQuota)}");
        }

        if (request.ProhibitSendReceiveQuota != null)
        {
            setMailboxParams.Add($"-ProhibitSendReceiveQuota {FormatNullableString(request.ProhibitSendReceiveQuota)}");
        }

        if (request.MaxSendSize != null)
        {
            setMailboxParams.Add($"-MaxSendSize {FormatNullableString(request.MaxSendSize)}");
        }

        if (request.MaxReceiveSize != null)
        {
            setMailboxParams.Add($"-MaxReceiveSize {FormatNullableString(request.MaxReceiveSize)}");
        }

        if (setMailboxParams.Count > 0)
        {
            scriptBuilder.AppendLine($"Set-Mailbox -Identity '{escapedIdentity}' {string.Join(" ", setMailboxParams)}");
        }

        var setCasMailboxParams = new List<string>();

        if (request.OwaEnabled.HasValue && CanIncludeCasParameter(capabilities, nameof(FeatureCapabilitiesDto.CanSetCasOwaEnabled)))
        {
            setCasMailboxParams.Add($"-OWAEnabled {ToPsBoolLiteral(request.OwaEnabled.Value)}");
        }

        if (request.ActiveSyncEnabled.HasValue && CanIncludeCasParameter(capabilities, nameof(FeatureCapabilitiesDto.CanSetCasActiveSyncEnabled)))
        {
            setCasMailboxParams.Add($"-ActiveSyncEnabled {ToPsBoolLiteral(request.ActiveSyncEnabled.Value)}");
        }

        if (request.MapiEnabled.HasValue && CanIncludeCasParameter(capabilities, nameof(FeatureCapabilitiesDto.CanSetCasMapiEnabled)))
        {
            setCasMailboxParams.Add($"-MAPIEnabled {ToPsBoolLiteral(request.MapiEnabled.Value)}");
        }

        if (request.PopEnabled.HasValue && CanIncludeCasParameter(capabilities, nameof(FeatureCapabilitiesDto.CanSetCasPopEnabled)))
        {
            setCasMailboxParams.Add($"-PopEnabled {ToPsBoolLiteral(request.PopEnabled.Value)}");
        }

        if (request.ImapEnabled.HasValue && CanIncludeCasParameter(capabilities, nameof(FeatureCapabilitiesDto.CanSetCasImapEnabled)))
        {
            setCasMailboxParams.Add($"-ImapEnabled {ToPsBoolLiteral(request.ImapEnabled.Value)}");
        }

        if (request.SmtpClientAuthenticationDisabled.HasValue &&
            CanIncludeCasParameter(capabilities, nameof(FeatureCapabilitiesDto.CanSetCasSmtpClientAuthenticationDisabled)))
        {
            setCasMailboxParams.Add($"-SmtpClientAuthenticationDisabled {ToPsBoolLiteral(request.SmtpClientAuthenticationDisabled.Value)}");
        }

        if (setCasMailboxParams.Count > 0)
        {
            scriptBuilder.AppendLine($"Set-CASMailbox -Identity '{escapedIdentity}' {string.Join(" ", setCasMailboxParams)}");
        }

        return scriptBuilder.ToString();
    }

    internal static string BuildSetRetentionPolicyScript(SetRetentionPolicyRequest request)
    {
        var escapedIdentity = EscapePs(request.Identity);
        var policyValue = FormatNullableString(request.PolicyName);
        return $"Set-Mailbox -Identity '{escapedIdentity}' -RetentionPolicy {policyValue}";
    }

    internal static string BuildSetMailboxAutoReplyConfigurationScript(SetMailboxAutoReplyConfigurationRequest request)
    {
        var escapedIdentity = EscapePs(request.Identity);
        var scriptBuilder = new StringBuilder();

        scriptBuilder.AppendLine("$params = @{}");
        scriptBuilder.AppendLine($"$params.Identity = '{escapedIdentity}'");
        scriptBuilder.AppendLine($"$params.AutoReplyState = '{request.AutoReplyState}'");

        if (request.StartTime.HasValue)
        {
            scriptBuilder.AppendLine($"$params.StartTime = [datetime]::Parse('{request.StartTime.Value:o}')");
        }

        if (request.EndTime.HasValue)
        {
            scriptBuilder.AppendLine($"$params.EndTime = [datetime]::Parse('{request.EndTime.Value:o}')");
        }

        if (request.InternalMessage != null)
        {
            scriptBuilder.AppendLine($"$params.InternalMessage = {FormatNullableMessage(request.InternalMessage)}");
        }

        if (request.ExternalMessage != null)
        {
            scriptBuilder.AppendLine($"$params.ExternalMessage = {FormatNullableMessage(request.ExternalMessage)}");
        }

        if (!string.IsNullOrWhiteSpace(request.ExternalAudience))
        {
            scriptBuilder.AppendLine($"$params.ExternalAudience = '{EscapePs(request.ExternalAudience)}'");
        }

        scriptBuilder.AppendLine("Set-MailboxAutoReplyConfiguration @params");
        return scriptBuilder.ToString();
    }

    internal static (string Script, Dictionary<string, object>? Parameters) BuildCreateMailboxCommand(CreateMailboxRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.DisplayName) ||
            string.IsNullOrWhiteSpace(request.Alias) ||
            string.IsNullOrWhiteSpace(request.PrimarySmtpAddress))
        {
            throw new InvalidOperationException("DisplayName, Alias and PrimarySmtpAddress are required to create a mailbox.");
        }

        var escapedDisplayName = EscapePs(request.DisplayName);
        var escapedAlias = EscapePs(request.Alias);
        var escapedPrimarySmtpAddress = EscapePs(request.PrimarySmtpAddress);
        var mailboxType = request.MailboxType?.Trim();

        if (string.Equals(mailboxType, "User", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(mailboxType, "Regular", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(mailboxType, "UserMailbox", StringComparison.OrdinalIgnoreCase))
        {
            if (string.IsNullOrWhiteSpace(request.Password))
            {
                throw new InvalidOperationException("Password is required to create a user mailbox.");
            }

            var script = $@"
$securePassword = ConvertTo-SecureString $PlainTextPassword -AsPlainText -Force
New-Mailbox -Name '{escapedDisplayName}' -DisplayName '{escapedDisplayName}' -Alias '{escapedAlias}' -MicrosoftOnlineServicesID '{escapedPrimarySmtpAddress}' -UserPrincipalName '{escapedPrimarySmtpAddress}' -PrimarySmtpAddress '{escapedPrimarySmtpAddress}' -Password $securePassword
";

            return (script, new Dictionary<string, object>
            {
                ["PlainTextPassword"] = request.Password
            });
        }

        return ($"New-Mailbox -Shared -Name '{escapedDisplayName}' -DisplayName '{escapedDisplayName}' -Alias '{escapedAlias}' -PrimarySmtpAddress '{escapedPrimarySmtpAddress}'", null);
    }

    internal static string BuildConvertMailboxTypeScript(string identity, string mailboxType)
        => $"Set-Mailbox -Identity '{EscapePs(identity)}' -Type {mailboxType}";

    internal static (string Script, Dictionary<string, object> Parameters) BuildRestoreMailboxCommand(RestoreMailboxRequest request)
    {
        var script = @"
param(
    [string]$SourceIdentity,
    [string]$TargetMailbox,
    [bool]$AllowLegacyDnMismatch
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($TargetMailbox)) {
    $TargetMailbox = $null
}

$result = [ordered]@{
    Success = $false
    Scenario = 'Unknown'
    Action = $null
    Status = 'NotStarted'
    StatusDetail = $null
    PercentComplete = $null
    RequestGuid = $null
    ErrorCode = $null
    ErrorMessage = $null
    SourceIdentity = $SourceIdentity
    TargetMailbox = $TargetMailbox
}

function Resolve-ErrorCode([string]$message) {
    if ($message -match 'TargetMailboxRequired') { return 'TargetMailboxRequired' }
    if ($message -match 'LegacyDN|legacy dn') { return 'LegacyDnMismatch' }
    if ($message -match 'Access is denied|not authorized|insufficient permissions|permission') { return 'PermissionDenied' }
    if ($message -match 'Cannot find|couldn''t be found|was not found|object with identity') { return 'UserNotFound' }
    return 'Unknown'
}

function Set-ErrorResult([string]$code, [string]$message) {
    $result.Success = $false
    $result.Status = 'Failed'
    $result.ErrorCode = $code
    $result.ErrorMessage = $message
    return $result
}

try {
    Write-Verbose ""Resolving mailbox state for $SourceIdentity...""
    $mailbox = $null

    try { $mailbox = Get-Mailbox -Identity $SourceIdentity -ErrorAction Stop } catch { $mailbox = $null }

    if ($mailbox) {
        $result.Scenario = 'Existing'
    } else {
        try { $mailbox = Get-Mailbox -SoftDeletedMailbox -Identity $SourceIdentity -ErrorAction Stop } catch { $mailbox = $null }

        if ($mailbox) {
            $result.Scenario = 'SoftDeleted'
        } else {
            try { $mailbox = Get-Mailbox -InactiveMailboxOnly -Identity $SourceIdentity -ErrorAction Stop } catch { $mailbox = $null }

            if ($mailbox) {
                $result.Scenario = 'Inactive'
            } else {
                $result.Scenario = 'NotFound'
                return Set-ErrorResult 'UserNotFound' ""Mailbox not found for identity $SourceIdentity""
            }
        }
    }

    if ($result.Scenario -eq 'SoftDeleted') {
        Write-Verbose 'Restoring soft-deleted mailbox via Undo-SoftDeletedMailbox...'
        $result.Action = 'Undo-SoftDeletedMailbox'
        $identity = if ($mailbox.Guid) { $mailbox.Guid } else { $SourceIdentity }
        if ($AllowLegacyDnMismatch) {
            Undo-SoftDeletedMailbox -Identity $identity -AllowLegacyDNMismatch -Confirm:$false -ErrorAction Stop
        } else {
            Undo-SoftDeletedMailbox -Identity $identity -Confirm:$false -ErrorAction Stop
        }

        $result.Success = $true
        $result.Status = 'Completed'
        $result.StatusDetail = 'Soft-deleted mailbox restored'
        $result.PercentComplete = 100
        return $result
    }

    if (-not $TargetMailbox) {
        throw 'TargetMailboxRequired'
    }

    Write-Verbose 'Creating mailbox restore request...'
    $result.Action = 'New-MailboxRestoreRequest'
    $params = @{
        SourceMailbox = if ($mailbox.Guid) { $mailbox.Guid } else { $SourceIdentity }
        TargetMailbox = $TargetMailbox
        ErrorAction = 'Stop'
    }

    if ($AllowLegacyDnMismatch) {
        $params['AllowLegacyDNMismatch'] = $true
    }

    $restoreRequest = New-MailboxRestoreRequest @params
    $result.RequestGuid = if ($restoreRequest.RequestGuid) { $restoreRequest.RequestGuid.ToString() } else { $null }
    $result.Status = 'InProgress'

    try {
        $stats = Get-MailboxRestoreRequestStatistics -Identity $restoreRequest.Identity -ErrorAction Stop
        if ($stats.Status) { $result.Status = $stats.Status.ToString() }
        if ($stats.StatusDetail) { $result.StatusDetail = $stats.StatusDetail.ToString() }
        if ($null -ne $stats.PercentComplete) { $result.PercentComplete = [int]$stats.PercentComplete }
    } catch {
        $result.StatusDetail = $_.Exception.Message
    }

    if (-not $result.StatusDetail) {
        if ($result.Scenario -eq 'Inactive') {
            $result.StatusDetail = 'Inactive mailbox restore request submitted'
        } else {
            $result.StatusDetail = 'Mailbox restore request submitted'
        }
    }

    $result.Success = $true
    return $result
}
catch {
    $message = $_.Exception.Message
    $code = Resolve-ErrorCode $message
    return Set-ErrorResult $code $message
}";

        return (script, new Dictionary<string, object>
        {
            ["SourceIdentity"] = request.SourceIdentity,
            ["TargetMailbox"] = request.TargetMailbox ?? string.Empty,
            ["AllowLegacyDnMismatch"] = request.AllowLegacyDnMismatch
        });
    }

    internal static string FormatNullableString(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "$null";
        }

        return $"'{EscapePs(value)}'";
    }

    internal static string FormatNullableMessage(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var base64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(value));
        return $"[System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('{base64}'))";
    }

    private static List<string> BuildMailboxEmailAddresses(string primarySmtpAddress, IEnumerable<string>? proxyAddresses)
    {
        var emailAddresses = new List<string> { $"SMTP:{primarySmtpAddress}" };
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            $"SMTP:{primarySmtpAddress}"
        };

        if (proxyAddresses == null)
        {
            return emailAddresses;
        }

        foreach (var proxyAddress in proxyAddresses)
        {
            var normalized = NormalizeProxyAddress(proxyAddress);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                continue;
            }

            if (TryGetSmtpAddress(normalized, out var smtpAddress) &&
                string.Equals(smtpAddress, primarySmtpAddress, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (seen.Add(normalized))
            {
                emailAddresses.Add(normalized);
            }
        }

        return emailAddresses;
    }

    private static string NormalizePrimarySmtpAddress(string? value)
        => string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();

    private static string NormalizeProxyAddress(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var trimmed = value.Trim();
        var separatorIndex = trimmed.IndexOf(':');
        if (separatorIndex <= 0)
        {
            return $"smtp:{trimmed}";
        }

        var prefix = trimmed[..separatorIndex].Trim();
        var address = trimmed[(separatorIndex + 1)..].Trim();
        if (string.IsNullOrWhiteSpace(address))
        {
            return string.Empty;
        }

        return string.Equals(prefix, "smtp", StringComparison.OrdinalIgnoreCase)
            ? $"smtp:{address}"
            : $"{prefix}:{address}";
    }

    private static bool TryGetSmtpAddress(string proxyAddress, out string smtpAddress)
    {
        smtpAddress = string.Empty;
        var separatorIndex = proxyAddress.IndexOf(':');
        if (separatorIndex <= 0)
        {
            return false;
        }

        var prefix = proxyAddress[..separatorIndex];
        if (!string.Equals(prefix, "smtp", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        smtpAddress = proxyAddress[(separatorIndex + 1)..];
        return !string.IsNullOrWhiteSpace(smtpAddress);
    }

    private static string EscapePs(string? value)
        => (value ?? string.Empty).Replace("'", "''");

    private static bool CanIncludeCasParameter(FeatureCapabilitiesDto? capabilities, string propertyName)
    {
        if (capabilities == null)
        {
            return true;
        }

        return propertyName switch
        {
            nameof(FeatureCapabilitiesDto.CanSetCasOwaEnabled) => capabilities.CanSetCasOwaEnabled,
            nameof(FeatureCapabilitiesDto.CanSetCasActiveSyncEnabled) => capabilities.CanSetCasActiveSyncEnabled,
            nameof(FeatureCapabilitiesDto.CanSetCasMapiEnabled) => capabilities.CanSetCasMapiEnabled,
            nameof(FeatureCapabilitiesDto.CanSetCasPopEnabled) => capabilities.CanSetCasPopEnabled,
            nameof(FeatureCapabilitiesDto.CanSetCasImapEnabled) => capabilities.CanSetCasImapEnabled,
            nameof(FeatureCapabilitiesDto.CanSetCasSmtpClientAuthenticationDisabled) => capabilities.CanSetCasSmtpClientAuthenticationDisabled,
            _ => false
        };
    }

    private static string ToPsBoolLiteral(bool value)
        => value ? "$true" : "$false";
}

