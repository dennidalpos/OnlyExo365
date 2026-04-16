using ExchangeAdmin.Contracts.Dtos;
using System.Management.Automation;

namespace ExchangeAdmin.Worker.PowerShell;

internal sealed partial class ExoMailboxReportingCommands : ExoCommandModuleBase
{
    public ExoMailboxReportingCommands(PowerShellEngine engine)
        : base(engine)
    {
    }

    public async Task<GetMailboxSpaceReportResponse> GetMailboxSpaceReportAsync(
        GetMailboxSpaceReportRequest request,
        Action<string, string>? onLog,
        Action<int, int>? onProgress,
        CancellationToken cancellationToken)
    {
        var script = @"
$ErrorActionPreference = 'Stop'

function Get-BytesFromSize($size) {
    if ($null -eq $size) { return $null }
    $text = $size.ToString()
    if ($text -match '\(([^)]+)\s+byte[s]?\)') {
        $numeric = ($Matches[1] -replace '[^\d]', '')
        if (-not [string]::IsNullOrWhiteSpace($numeric)) {
            return [long]$numeric
        }
    }
    return $null
}

$statisticsFailures = New-Object System.Collections.Generic.List[string]
$mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox -ErrorAction Stop
$total = @($mailboxes).Count
$index = 0

foreach ($mbx in $mailboxes) {
    $index++
    $stats = $null
    try {
        $stats = Get-MailboxStatistics -Identity $mbx.Identity -ErrorAction Stop
    } catch {
        $stats = $null
        $identity = if ($mbx.PrimarySmtpAddress) { $mbx.PrimarySmtpAddress.ToString() } else { $mbx.Identity.ToString() }
        $statisticsFailures.Add($identity)
    }

    [PSCustomObject]@{
        EntryType = 'Mailbox'
        Index = $index
        TotalCount = $total
        Identity = $mbx.Identity.ToString()
        DisplayName = $mbx.DisplayName
        PrimarySmtpAddress = $mbx.PrimarySmtpAddress.ToString()
        TotalItemSize = if ($stats -and $stats.TotalItemSize) { $stats.TotalItemSize.ToString() } else { $null }
        TotalItemSizeBytes = Get-BytesFromSize $stats.TotalItemSize
        ProhibitSendQuota = if ($mbx.ProhibitSendQuota) { $mbx.ProhibitSendQuota.ToString() } else { $null }
        ProhibitSendQuotaBytes = Get-BytesFromSize $mbx.ProhibitSendQuota
        ProhibitSendReceiveQuota = if ($mbx.ProhibitSendReceiveQuota) { $mbx.ProhibitSendReceiveQuota.ToString() } else { $null }
        ProhibitSendReceiveQuotaBytes = Get-BytesFromSize $mbx.ProhibitSendReceiveQuota
        IssueWarningQuota = if ($mbx.IssueWarningQuota) { $mbx.IssueWarningQuota.ToString() } else { $null }
        IssueWarningQuotaBytes = Get-BytesFromSize $mbx.IssueWarningQuota
    }
}

if ($statisticsFailures.Count -gt 0) {
    $warningPayload = @{
        Code = 'MailboxStatisticsUnavailable'
        Scope = 'MailboxSpaceReport'
        Message = ""Get-MailboxStatistics is not available for $($statisticsFailures.Count) mailbox(es). The report contains partial data.""
        AffectedItemCount = $statisticsFailures.Count
        SampleItems = @($statisticsFailures | Select-Object -First 5)
        IsPartialData = $true
    } | ConvertTo-Json -Compress -Depth 4
    Write-Warning ""__EA_WARN__$warningPayload""
}

[PSCustomObject]@{
    EntryType = 'Summary'
    ProcessedMailboxCount = $total
}
";

        onLog?.Invoke("Information", "Generating mailbox space report...");

        var result = await Engine.ExecuteAsync(
            script,
            onVerbose: onLog,
            onOutput: output =>
            {
                if (onProgress == null || output == null)
                {
                    return;
                }

                if (TryReadProgress(output, out var current, out var total))
                {
                    onProgress(current, total);
                }
            },
            cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to get mailbox space report: {result.ErrorMessage}");
        }

        var warningDetails = ParseStructuredWarnings(result.Warning);
        var response = new GetMailboxSpaceReportResponse
        {
            WarningDetails = warningDetails,
            Warnings = ExtractWarningMessages(warningDetails),
            HasPartialData = warningDetails.Any(static warning => warning.IsPartialData)
        };

        foreach (var output in result.Output)
        {
            if (output == null)
            {
                continue;
            }

            if (TryReadProcessedMailboxCount(output) is { } processedCount)
            {
                response.ProcessedMailboxCount = processedCount;
                continue;
            }

            if (TryParseMailboxSpaceItem(output) is { } item)
            {
                response.Mailboxes.Add(item);
            }
        }

        if (response.ProcessedMailboxCount == 0)
        {
            response.ProcessedMailboxCount = response.Mailboxes.Count;
        }

        response.FailedMailboxCount = warningDetails
            .Where(static warning => string.Equals(warning.Code, "MailboxStatisticsUnavailable", StringComparison.Ordinal))
            .Select(static warning => warning.AffectedItemCount ?? 0)
            .FirstOrDefault();

        onLog?.Invoke("Information", $"Mailbox space report generated: {response.Mailboxes.Count} entries");

        return response;
    }

    public async Task<GetMailboxAccessReportResponse> GetMailboxAccessReportAsync(
        GetMailboxAccessReportRequest request,
        Action<string, string>? onLog,
        Action<int, int>? onProgress,
        CancellationToken cancellationToken)
    {
        var script = @"
$ErrorActionPreference = 'Stop'

$principalCache = @{}
$principalResolutionFallbacks = New-Object System.Collections.Generic.HashSet[string]
$fullAccessFailures = New-Object System.Collections.Generic.List[string]
$sendAsFailures = New-Object System.Collections.Generic.List[string]

function Is-SkippablePrincipal($value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return $true }
    if ($value -like 'NT AUTHORITY\*') { return $true }
    if ($value -like 'S-1-5-*') { return $true }
    if ($value -eq 'SELF') { return $true }
    return $false
}

function Resolve-PrincipalIdentity($value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return $value }
    if ($principalCache.ContainsKey($value)) { return $principalCache[$value] }
    $resolved = $value
    try {
        $recipient = Get-Recipient -Identity $value -ErrorAction Stop
        if ($null -ne $recipient) {
            if ($recipient.PrimarySmtpAddress) { $resolved = $recipient.PrimarySmtpAddress.ToString() }
            elseif ($recipient.WindowsEmailAddress) { $resolved = $recipient.WindowsEmailAddress.ToString() }
            elseif ($recipient.UserPrincipalName) { $resolved = $recipient.UserPrincipalName.ToString() }
            elseif ($recipient.Name) { $resolved = $recipient.Name.ToString() }
        }
    } catch {
        [void]$principalResolutionFallbacks.Add($value)
    }
    $principalCache[$value] = $resolved
    return $resolved
}

$mailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox -Properties GrantSendOnBehalfTo -ErrorAction Stop
$total = @($mailboxes).Count
$index = 0

foreach ($mbx in $mailboxes) {
    $index++
    $mailboxIdentity = $mbx.Identity.ToString()
    $mailboxDisplayName = $mbx.DisplayName
    $mailboxPrimarySmtp = if ($mbx.PrimarySmtpAddress) { $mbx.PrimarySmtpAddress.ToString() } else { '' }

    try {
        $fullAccess = Get-MailboxPermission -Identity $mailboxIdentity -ErrorAction Stop |
            Where-Object { $_.IsInherited -eq $false -and $_.Deny -eq $false -and -not (Is-SkippablePrincipal $_.User.ToString()) }
        foreach ($perm in $fullAccess) {
            $resolvedUser = Resolve-PrincipalIdentity $perm.User.ToString()
            if (Is-SkippablePrincipal $resolvedUser) { continue }
            [PSCustomObject]@{
                EntryType = 'Grant'
                Index = $index
                TotalCount = $total
                User = $resolvedUser
                MailboxIdentity = $mailboxIdentity
                MailboxDisplayName = $mailboxDisplayName
                MailboxPrimarySmtpAddress = $mailboxPrimarySmtp
                PermissionType = 'FullAccess'
                AccessRights = @($perm.AccessRights | ForEach-Object { $_.ToString() })
            }
        }
    } catch {
        $fullAccessFailures.Add($mailboxPrimarySmtp)
    }

    try {
        $sendAs = Get-RecipientPermission -Identity $mailboxIdentity -ErrorAction Stop |
            Where-Object { $_.IsInherited -eq $false -and -not (Is-SkippablePrincipal $_.Trustee.ToString()) }
        foreach ($perm in $sendAs) {
            $resolvedUser = Resolve-PrincipalIdentity $perm.Trustee.ToString()
            if (Is-SkippablePrincipal $resolvedUser) { continue }
            [PSCustomObject]@{
                EntryType = 'Grant'
                Index = $index
                TotalCount = $total
                User = $resolvedUser
                MailboxIdentity = $mailboxIdentity
                MailboxDisplayName = $mailboxDisplayName
                MailboxPrimarySmtpAddress = $mailboxPrimarySmtp
                PermissionType = 'SendAs'
                AccessRights = @($perm.AccessRights | ForEach-Object { $_.ToString() })
            }
        }
    } catch {
        $sendAsFailures.Add($mailboxPrimarySmtp)
    }

    foreach ($delegate in @($mbx.GrantSendOnBehalfTo)) {
        if ($null -eq $delegate) { continue }
        $delegateValue = $delegate.ToString()
        if (Is-SkippablePrincipal $delegateValue) { continue }
        $resolvedUser = Resolve-PrincipalIdentity $delegateValue
        if (Is-SkippablePrincipal $resolvedUser) { continue }
        [PSCustomObject]@{
            EntryType = 'Grant'
            Index = $index
            TotalCount = $total
            User = $resolvedUser
            MailboxIdentity = $mailboxIdentity
            MailboxDisplayName = $mailboxDisplayName
            MailboxPrimarySmtpAddress = $mailboxPrimarySmtp
            PermissionType = 'SendOnBehalf'
            AccessRights = @('SendOnBehalf')
        }
    }
}

if ($fullAccessFailures.Count -gt 0) {
    $warningPayload = @{
        Code = 'FullAccessQueryFailed'
        Scope = 'MailboxAccessReport'
        Message = ""Get-MailboxPermission failed for $($fullAccessFailures.Count) mailbox(es). The report is partial.""
        AffectedItemCount = $fullAccessFailures.Count
        SampleItems = @($fullAccessFailures | Select-Object -First 5)
        IsPartialData = $true
    } | ConvertTo-Json -Compress -Depth 4
    Write-Warning ""__EA_WARN__$warningPayload""
}

if ($sendAsFailures.Count -gt 0) {
    $warningPayload = @{
        Code = 'SendAsQueryFailed'
        Scope = 'MailboxAccessReport'
        Message = ""Get-RecipientPermission failed for $($sendAsFailures.Count) mailbox(es). The report is partial.""
        AffectedItemCount = $sendAsFailures.Count
        SampleItems = @($sendAsFailures | Select-Object -First 5)
        IsPartialData = $true
    } | ConvertTo-Json -Compress -Depth 4
    Write-Warning ""__EA_WARN__$warningPayload""
}

if ($principalResolutionFallbacks.Count -gt 0) {
    $warningPayload = @{
        Code = 'PrincipalResolutionFallback'
        Scope = 'MailboxAccessReport'
        Message = ""Some identities were not resolved through Get-Recipient; the report uses the raw value for $($principalResolutionFallbacks.Count) principal(s).""
        AffectedItemCount = $principalResolutionFallbacks.Count
        SampleItems = @($principalResolutionFallbacks | Select-Object -First 5)
        IsPartialData = $false
    } | ConvertTo-Json -Compress -Depth 4
    Write-Warning ""__EA_WARN__$warningPayload""
}

[PSCustomObject]@{
    EntryType = 'Summary'
    ProcessedMailboxCount = $total
}
";

        onLog?.Invoke("Information", "Generating mailbox access report...");

        var result = await Engine.ExecuteAsync(
            script,
            onVerbose: onLog,
            onOutput: output =>
            {
                if (onProgress == null || output == null)
                {
                    return;
                }

                if (TryReadProgress(output, out var current, out var total))
                {
                    onProgress(current, total);
                }
            },
            cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to get mailbox access report: {result.ErrorMessage}");
        }

        var warningDetails = ParseStructuredWarnings(result.Warning);
        var response = new GetMailboxAccessReportResponse
        {
            WarningDetails = warningDetails,
            Warnings = ExtractWarningMessages(warningDetails),
            HasPartialData = warningDetails.Any(static warning => warning.IsPartialData)
        };

        foreach (var output in result.Output)
        {
            if (output == null)
            {
                continue;
            }

            if (TryReadProcessedMailboxCount(output) is { } processedCount)
            {
                response.ProcessedMailboxCount = processedCount;
                continue;
            }

            if (TryParseMailboxAccessGrant(output) is not { } grant)
            {
                continue;
            }

            response.Grants.Add(grant);
        }

        response.FailedMailboxCount = warningDetails
            .Where(static warning => warning.IsPartialData)
            .Sum(static warning => warning.AffectedItemCount ?? 0);

        onLog?.Invoke("Information", $"Mailbox access report generated: {response.Grants.Count} grants");

        return response;
    }

    internal static bool TryReadProgress(PSObject output, out int current, out int total)
    {
        current = GetNullableInt(output, "Index") ?? 0;
        total = GetNullableInt(output, "TotalCount") ?? 0;
        return current > 0 && total > 0;
    }

    internal static int? TryReadProcessedMailboxCount(PSObject output)
    {
        return string.Equals(GetNullableString(output, "EntryType"), "Summary", StringComparison.Ordinal)
            ? GetNullableInt(output, "ProcessedMailboxCount")
            : null;
    }

    internal static MailboxSpaceItemDto? TryParseMailboxSpaceItem(PSObject output)
    {
        if (!string.Equals(GetNullableString(output, "EntryType"), "Mailbox", StringComparison.Ordinal))
        {
            return null;
        }

        return new MailboxSpaceItemDto
        {
            Identity = GetString(output, "Identity"),
            DisplayName = GetString(output, "DisplayName"),
            PrimarySmtpAddress = GetString(output, "PrimarySmtpAddress"),
            TotalItemSize = GetNullableString(output, "TotalItemSize"),
            TotalItemSizeBytes = GetNullableLong(output, "TotalItemSizeBytes"),
            ProhibitSendQuota = GetNullableString(output, "ProhibitSendQuota"),
            ProhibitSendQuotaBytes = GetNullableLong(output, "ProhibitSendQuotaBytes"),
            ProhibitSendReceiveQuota = GetNullableString(output, "ProhibitSendReceiveQuota"),
            ProhibitSendReceiveQuotaBytes = GetNullableLong(output, "ProhibitSendReceiveQuotaBytes"),
            IssueWarningQuota = GetNullableString(output, "IssueWarningQuota"),
            IssueWarningQuotaBytes = GetNullableLong(output, "IssueWarningQuotaBytes")
        };
    }

    internal static MailboxAccessGrantDto? TryParseMailboxAccessGrant(PSObject output)
    {
        if (!string.Equals(GetNullableString(output, "EntryType"), "Grant", StringComparison.Ordinal))
        {
            return null;
        }

        var user = GetNullableString(output, "User");
        if (string.IsNullOrWhiteSpace(user))
        {
            return null;
        }

        return new MailboxAccessGrantDto
        {
            User = user,
            MailboxIdentity = GetString(output, "MailboxIdentity"),
            MailboxDisplayName = GetString(output, "MailboxDisplayName"),
            MailboxPrimarySmtpAddress = GetString(output, "MailboxPrimarySmtpAddress"),
            PermissionType = GetString(output, "PermissionType"),
            AccessRights = ConvertToStringList(GetPropertyValue(output, "AccessRights"))
        };
    }

    public async Task<MailboxPermissionsDto> GetMailboxPermissionsAsync(
        string identity,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var permissions = new MailboxPermissionsDto();
        var escapedIdentity = identity.Replace("'", "''");

        var fullAccessScript = $@"
try {{
    $perms = Get-MailboxPermission -Identity '{escapedIdentity}' -ErrorAction Stop |
        Where-Object {{ $_.User -notlike 'NT AUTHORITY\*' -and $_.User -notlike 'S-1-5-*' -and $_.IsInherited -eq $false }}
    @($perms | ForEach-Object {{
        @{{
            Identity = $_.Identity.ToString()
            User = $_.User.ToString()
            AccessRights = @($_.AccessRights | ForEach-Object {{ $_.ToString() }})
            IsInherited = $_.IsInherited
            Deny = $_.Deny
            InheritanceType = $_.InheritanceType.ToString()
        }}
    }})
}}
catch {{
    @()
}}
";

        var fullAccessResult = await Engine.ExecuteAsync(fullAccessScript, onVerbose: onLog, cancellationToken: cancellationToken);
        if (fullAccessResult.Success)
        {
            foreach (var output in fullAccessResult.Output)
            {
                if (output.BaseObject is System.Collections.Hashtable hash)
                {
                    permissions.FullAccessPermissions.Add(new MailboxPermissionEntryDto
                    {
                        Identity = hash["Identity"]?.ToString() ?? "",
                        User = hash["User"]?.ToString() ?? "",
                        AccessRights = ConvertToStringList(hash["AccessRights"]),
                        IsInherited = hash["IsInherited"] as bool? ?? false,
                        Deny = hash["Deny"] as bool? ?? false,
                        InheritanceType = hash["InheritanceType"]?.ToString()
                    });
                }
            }
        }

        cancellationToken.ThrowIfCancellationRequested();

        var sendAsScript = $@"
try {{
    $perms = Get-RecipientPermission -Identity '{escapedIdentity}' -ErrorAction Stop |
        Where-Object {{ $_.Trustee -notlike 'NT AUTHORITY\*' -and $_.Trustee -notlike 'S-1-5-*' }}
    @($perms | ForEach-Object {{
        $trustee = $_.Trustee.ToString()
        $displayName = $trustee
        $trusteeIdentity = $trustee
        try {{
            $recipient = Get-Recipient -Identity $trustee -ErrorAction Stop
            if ($recipient.DisplayName) {{ $displayName = $recipient.DisplayName }}
            if ($recipient.PrimarySmtpAddress) {{ $trusteeIdentity = $recipient.PrimarySmtpAddress.ToString() }}
            elseif ($recipient.ExternalDirectoryObjectId) {{ $trusteeIdentity = $recipient.ExternalDirectoryObjectId.ToString() }}
            elseif ($recipient.Identity) {{ $trusteeIdentity = $recipient.Identity.ToString() }}
        }} catch {{}}
        @{{
            Identity = $_.Identity.ToString()
            Trustee = $trustee
            ResolvedTrustee = $trusteeIdentity
            DisplayName = $displayName
            AccessControlType = $_.AccessControlType.ToString()
            AccessRights = @($_.AccessRights | ForEach-Object {{ $_.ToString() }})
            IsInherited = $_.IsInherited
        }}
    }})
}}
catch {{
    @()
}}
";

        var sendAsResult = await Engine.ExecuteAsync(sendAsScript, onVerbose: onLog, cancellationToken: cancellationToken);
        if (sendAsResult.Success)
        {
            foreach (var output in sendAsResult.Output)
            {
                if (output.BaseObject is System.Collections.Hashtable hash)
                {
                    permissions.SendAsPermissions.Add(new RecipientPermissionEntryDto
                    {
                        Identity = hash["Identity"]?.ToString() ?? "",
                        Trustee = hash["Trustee"]?.ToString() ?? "",
                        ResolvedTrustee = hash["ResolvedTrustee"]?.ToString() ?? "",
                        DisplayName = hash["DisplayName"]?.ToString() ?? "",
                        AccessControlType = hash["AccessControlType"]?.ToString() ?? "",
                        AccessRights = ConvertToStringList(hash["AccessRights"]),
                        IsInherited = hash["IsInherited"] as bool? ?? false
                    });
                }
            }
        }

        cancellationToken.ThrowIfCancellationRequested();

        var sendOnBehalfScript = $@"
try {{
    $mbx = Get-Mailbox -Identity '{escapedIdentity}' -ErrorAction Stop
    @($mbx.GrantSendOnBehalfTo | ForEach-Object {{
        $rawIdentity = $_.ToString()
        $displayName = $null
        try {{
            $recipient = Get-Recipient -Identity $rawIdentity -ErrorAction Stop
            if ($recipient.DisplayName) {{ $displayName = $recipient.DisplayName }}
            elseif ($recipient.PrimarySmtpAddress) {{ $displayName = $recipient.PrimarySmtpAddress.ToString() }}
        }} catch {{}}
        @{{
            Identity = $rawIdentity
            DisplayName = if ($displayName) {{ $displayName }} else {{ $rawIdentity }}
        }}
    }})
}}
catch {{
    @()
}}
";

        var sendOnBehalfResult = await Engine.ExecuteAsync(sendOnBehalfScript, onVerbose: onLog, cancellationToken: cancellationToken);
        if (sendOnBehalfResult.Success)
        {
            foreach (var output in sendOnBehalfResult.Output)
            {
                if (output.BaseObject is System.Collections.Hashtable hash)
                {
                    permissions.SendOnBehalfPermissions.Add(new SendOnBehalfPermissionEntryDto
                    {
                        Identity = hash["Identity"]?.ToString() ?? string.Empty,
                        DisplayName = hash["DisplayName"]?.ToString() ?? string.Empty
                    });
                }
            }
        }

        return permissions;
    }

    public async Task SetMailboxPermissionAsync(
        SetMailboxPermissionRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var escapedIdentity = request.Identity.Replace("'", "''");
        var escapedUser = request.User.Replace("'", "''");

        string script = request.PermissionType switch
        {
            PermissionType.FullAccess when request.Action == PermissionAction.Add =>
                $"Add-MailboxPermission -Identity '{escapedIdentity}' -User '{escapedUser}' -AccessRights FullAccess -AutoMapping ${request.AutoMapping ?? true} -Confirm:$false",
            PermissionType.FullAccess when request.Action == PermissionAction.Modify =>
                $@"
Remove-MailboxPermission -Identity '{escapedIdentity}' -User '{escapedUser}' -AccessRights FullAccess -Confirm:$false
Add-MailboxPermission -Identity '{escapedIdentity}' -User '{escapedUser}' -AccessRights FullAccess -AutoMapping ${request.AutoMapping ?? true} -Confirm:$false
",
            PermissionType.FullAccess =>
                $"Remove-MailboxPermission -Identity '{escapedIdentity}' -User '{escapedUser}' -AccessRights FullAccess -Confirm:$false",
            PermissionType.SendAs when request.Action == PermissionAction.Add =>
                $"Add-RecipientPermission -Identity '{escapedIdentity}' -Trustee '{escapedUser}' -AccessRights SendAs -Confirm:$false -ErrorAction Stop",
            PermissionType.SendAs =>
                $"Remove-RecipientPermission -Identity '{escapedIdentity}' -Trustee '{escapedUser}' -AccessRights SendAs -Confirm:$false -ErrorAction Stop",
            PermissionType.SendOnBehalf when request.Action == PermissionAction.Add =>
                $@"
$mbx = Get-Mailbox -Identity '{escapedIdentity}'
$current = @($mbx.GrantSendOnBehalfTo)
$current += '{escapedUser}'
Set-Mailbox -Identity '{escapedIdentity}' -GrantSendOnBehalfTo $current
",
            PermissionType.SendOnBehalf =>
                $@"
$mbx = Get-Mailbox -Identity '{escapedIdentity}'
$current = @($mbx.GrantSendOnBehalfTo) | Where-Object {{ $_.ToString() -ne '{escapedUser}' }}
Set-Mailbox -Identity '{escapedIdentity}' -GrantSendOnBehalfTo $current
",
            _ => throw new ArgumentException($"Unknown permission type: {request.PermissionType}")
        };

        var result = await Engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to {request.Action} {request.PermissionType} permission: {result.ErrorMessage}");
        }
    }

    public async Task<ApplyPermissionsDeltaPlanResponse> ApplyPermissionsDeltaPlanAsync(
        ApplyPermissionsDeltaPlanRequest request,
        Action<string, string>? onLog,
        Action<int, int>? onProgress,
        CancellationToken cancellationToken)
    {
        var response = new ApplyPermissionsDeltaPlanResponse
        {
            Identity = request.Identity,
            TotalActions = request.Actions.Count
        };

        for (var i = 0; i < request.Actions.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var action = request.Actions[i];
            var result = new PermissionActionResultDto { Action = action };

            try
            {
                await SetMailboxPermissionAsync(new SetMailboxPermissionRequest
                {
                    Identity = request.Identity,
                    User = action.User,
                    PermissionType = action.PermissionType,
                    Action = action.Action,
                    AutoMapping = action.AutoMapping
                }, onLog, cancellationToken);
                result.Success = true;
                response.SuccessfulActions++;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                response.FailedActions++;
            }

            response.Results.Add(result);
            onProgress?.Invoke(i + 1, request.Actions.Count);
        }

        return response;
    }

}
