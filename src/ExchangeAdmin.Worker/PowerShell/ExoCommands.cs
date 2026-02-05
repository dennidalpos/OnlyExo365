using System.Management.Automation;
using System.Text;
using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

public class ExoCommands
{
    private readonly PowerShellEngine _engine;
    private readonly CapabilityDetector _capabilityDetector;

    public ExoCommands(PowerShellEngine engine, CapabilityDetector capabilityDetector)
    {
        _engine = engine;
        _capabilityDetector = capabilityDetector;
    }

    public async Task<DashboardStatsDto> GetDashboardStatsAsync(
        GetDashboardStatsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var stats = new DashboardStatsDto();
        var warnings = new List<string>();
        onLog?.Invoke("Verbose", "Fetching mailbox counts...");

        var mailboxScript = @"
$counts = @{
    UserMailboxes = 0
    SharedMailboxes = 0
    RoomMailboxes = 0
    EquipmentMailboxes = 0
    IsApproximate = $false
}

try {
    $result = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox |
        Group-Object RecipientTypeDetails |
        Select-Object Name, Count

    foreach ($item in $result) {
        switch ($item.Name) {
            'UserMailbox' { $counts.UserMailboxes = $item.Count }
            'SharedMailbox' { $counts.SharedMailboxes = $item.Count }
            'RoomMailbox' { $counts.RoomMailboxes = $item.Count }
            'EquipmentMailbox' { $counts.EquipmentMailboxes = $item.Count }
        }
    }
}
catch {
    $counts.IsApproximate = $true
    try {
        $counts.UserMailboxes = (Get-Mailbox -ResultSize 1000 -RecipientTypeDetails UserMailbox).Count
        $counts.SharedMailboxes = (Get-Mailbox -ResultSize 1000 -RecipientTypeDetails SharedMailbox).Count
        $counts.RoomMailboxes = (Get-Mailbox -ResultSize 1000 -RecipientTypeDetails RoomMailbox).Count
        $counts.EquipmentMailboxes = (Get-Mailbox -ResultSize 1000 -RecipientTypeDetails EquipmentMailbox).Count
    }
    catch { }
}

$counts
";

        var mailboxResult = await _engine.ExecuteAsync(mailboxScript, onVerbose: onLog, cancellationToken: cancellationToken);

        if (mailboxResult.Success && mailboxResult.Output.Any())
        {
            var hash = mailboxResult.Output.First().BaseObject as System.Collections.Hashtable;
            if (hash != null)
            {
                stats.MailboxCounts = new MailboxCountsDto
                {
                    UserMailboxes = Convert.ToInt32(hash["UserMailboxes"] ?? 0),
                    SharedMailboxes = Convert.ToInt32(hash["SharedMailboxes"] ?? 0),
                    RoomMailboxes = Convert.ToInt32(hash["RoomMailboxes"] ?? 0),
                    EquipmentMailboxes = Convert.ToInt32(hash["EquipmentMailboxes"] ?? 0),
                    IsApproximate = hash["IsApproximate"] as bool? ?? false
                };
            }
        }

        cancellationToken.ThrowIfCancellationRequested();

        onLog?.Invoke("Verbose", "Fetching group counts...");

        var groupScript = @"
$counts = @{
    DistributionGroups = 0
    DynamicDistributionGroups = 0
}

try {
    $counts.DistributionGroups = (Get-DistributionGroup -ResultSize Unlimited).Count
}
catch {
    try { $counts.DistributionGroups = (Get-DistributionGroup -ResultSize 1000).Count } catch { }
}

try {
    $counts.DynamicDistributionGroups = (Get-DynamicDistributionGroup -ResultSize Unlimited).Count
}
catch {
    try { $counts.DynamicDistributionGroups = (Get-DynamicDistributionGroup -ResultSize 1000).Count } catch { }
}

$counts
";

        var groupResult = await _engine.ExecuteAsync(groupScript, onVerbose: onLog, cancellationToken: cancellationToken);

        if (groupResult.Success && groupResult.Output.Any())
        {
            var hash = groupResult.Output.First().BaseObject as System.Collections.Hashtable;
            if (hash != null)
            {
                stats.GroupCounts = new GroupCountsDto
                {
                    DistributionGroups = Convert.ToInt32(hash["DistributionGroups"] ?? 0),
                    DynamicDistributionGroups = Convert.ToInt32(hash["DynamicDistributionGroups"] ?? 0)
                };
            }
        }

        cancellationToken.ThrowIfCancellationRequested();

        var capabilities = _capabilityDetector.CachedCapabilities;
        if (request.IncludeUnifiedGroups && capabilities?.Features.CanGetUnifiedGroup == true)
        {
            onLog?.Invoke("Verbose", "Fetching unified group counts...");

            var unifiedScript = @"
try {
    $count = (Get-UnifiedGroup -ResultSize Unlimited).Count
    @{ Count = $count; Available = $true }
}
catch {
    @{ Count = 0; Available = $false; Error = $_.Exception.Message }
}
";
            var unifiedResult = await _engine.ExecuteAsync(unifiedScript, onVerbose: onLog, cancellationToken: cancellationToken);

            if (unifiedResult.Success && unifiedResult.Output.Any())
            {
                var hash = unifiedResult.Output.First().BaseObject as System.Collections.Hashtable;
                if (hash != null)
                {
                    stats.GroupCounts.UnifiedGroupsAvailable = hash["Available"] as bool? ?? false;
                    if (stats.GroupCounts.UnifiedGroupsAvailable)
                    {
                        stats.GroupCounts.UnifiedGroups = Convert.ToInt32(hash["Count"] ?? 0);
                    }
                    else
                    {
                        warnings.Add($"Unified Groups not available: {hash["Error"]}");
                    }
                }
            }
        }

        var totalMailboxes = stats.MailboxCounts.Total;
        if (totalMailboxes > stats.LargeTenantThreshold)
        {
            stats.IsLargeTenant = true;
            warnings.Add($"Large tenant detected ({totalMailboxes} mailboxes). Some operations may be slower.");
        }

        cancellationToken.ThrowIfCancellationRequested();

        if (stats.MailboxCounts.IsApproximate)
        {
            warnings.Add("Mailbox counts are approximate due to fallback limits.");
        }

        // Fetch tenant licenses (requires Microsoft Graph)
        onLog?.Invoke("Verbose", "Fetching tenant licenses...");
        try
        {
            stats.Licenses = await GetTenantLicensesAsync(cancellationToken);
        }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex)
        {
            onLog?.Invoke("Warning", $"Could not fetch tenant licenses: {ex.Message}");
        }

        cancellationToken.ThrowIfCancellationRequested();

        // Fetch admin role members (requires Microsoft Graph)
        onLog?.Invoke("Verbose", "Fetching admin role members...");
        try
        {
            stats.AdminUsers = await GetAdminRoleMembersAsync(cancellationToken);
        }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex)
        {
            onLog?.Invoke("Warning", $"Could not fetch admin role members: {ex.Message}");
        }

        stats.Warnings = warnings;
        stats.RetrievedAt = DateTime.UtcNow;

        return stats;
    }

    public async Task<GetMailboxesResponse> GetMailboxesAsync(
        GetMailboxesRequest request,
        Action<string, string>? onLog = null,
        Action<MailboxListItemDto>? onPartialOutput = null,
        CancellationToken cancellationToken = default)
    {
        var response = new GetMailboxesResponse
        {
            Skip = request.Skip,
            PageSize = request.PageSize,
            SearchQuery = request.SearchQuery
        };

        var filterParts = new List<string>();

        if (!string.IsNullOrWhiteSpace(request.RecipientTypeDetails))
        {
            filterParts.Add($"RecipientTypeDetails -eq '{request.RecipientTypeDetails}'");
        }

        if (!string.IsNullOrWhiteSpace(request.Filter))
        {
            filterParts.Add($"({request.Filter})");
        }

        var filterParam = filterParts.Count > 0 ? $"-Filter \"{string.Join(" -and ", filterParts)}\"" : "";

        var script = $@"
$allMailboxes = Get-Mailbox -ResultSize Unlimited {filterParam}
";

        if (!string.IsNullOrWhiteSpace(request.SearchQuery))
        {
            var escapedSearch = request.SearchQuery.Replace("'", "''");
            script += $@"
$allMailboxes = $allMailboxes | Where-Object {{
    $_.DisplayName -like '*{escapedSearch}*' -or
    $_.PrimarySmtpAddress -like '*{escapedSearch}*' -or
    $_.Alias -like '*{escapedSearch}*'
}}
";
        }

        var sortProperty = string.IsNullOrWhiteSpace(request.SortBy) ? "DisplayName" : request.SortBy;
        var sortDirection = request.SortDescending ? "-Descending" : "";
        script += $@"
$allMailboxes = $allMailboxes | Sort-Object {sortProperty} {sortDirection}
$totalCount = @($allMailboxes).Count

$pagedMailboxes = $allMailboxes | Select-Object -Skip {request.Skip} -First {request.PageSize}

@{{
    TotalCount = $totalCount
    Mailboxes = @($pagedMailboxes | ForEach-Object {{
        @{{
            Identity = $_.Identity.ToString()
            Guid = $_.ExchangeGuid.ToString()
            DisplayName = $_.DisplayName
            PrimarySmtpAddress = $_.PrimarySmtpAddress.ToString()
            RecipientType = $_.RecipientType.ToString()
            RecipientTypeDetails = $_.RecipientTypeDetails.ToString()
            Alias = $_.Alias
            IsInactiveMailbox = $_.IsInactiveMailbox
        }}
    }})
}}
";

        onLog?.Invoke("Verbose", $"Fetching mailboxes (skip={request.Skip}, pageSize={request.PageSize})...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (result.Success && result.Output.Any())
        {
            var firstOutput = result.Output.First();
            var hash = firstOutput?.BaseObject as System.Collections.Hashtable;
            if (hash != null)
            {
                response.TotalCount = Convert.ToInt32(hash["TotalCount"] ?? 0);

                var mailboxes = hash["Mailboxes"] as object[];
                if (mailboxes != null)
                {
                    foreach (var mbxObj in mailboxes)
                    {
                        if (mbxObj is System.Collections.Hashtable mbxHash)
                        {
                            var item = new MailboxListItemDto
                            {
                                Identity = mbxHash["Identity"]?.ToString() ?? "",
                                Guid = mbxHash["Guid"]?.ToString(),
                                DisplayName = mbxHash["DisplayName"]?.ToString() ?? "",
                                PrimarySmtpAddress = mbxHash["PrimarySmtpAddress"]?.ToString() ?? "",
                                RecipientType = mbxHash["RecipientType"]?.ToString() ?? "",
                                RecipientTypeDetails = mbxHash["RecipientTypeDetails"]?.ToString() ?? "",
                                Alias = mbxHash["Alias"]?.ToString(),
                                IsInactiveMailbox = mbxHash["IsInactiveMailbox"] as bool? ?? false
                            };

                            response.Mailboxes.Add(item);
                            onPartialOutput?.Invoke(item);
                        }
                    }
                }

                response.HasMore = (request.Skip + response.Mailboxes.Count) < response.TotalCount;
            }
        }

        onLog?.Invoke("Information", $"Retrieved {response.Mailboxes.Count} mailboxes (total: {response.TotalCount})");

        return response;
    }

    public async Task<GetDeletedMailboxesResponse> GetDeletedMailboxesAsync(
        GetDeletedMailboxesRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var response = new GetDeletedMailboxesResponse
        {
            Skip = request.Skip,
            PageSize = request.PageSize,
            SearchQuery = request.SearchQuery
        };

        var includeSoftDeleted = request.IncludeSoftDeleted ? "$true" : "$false";
        var includeInactive = request.IncludeInactive ? "$true" : "$false";
        var escapedSearch = request.SearchQuery?.Replace("'", "''") ?? string.Empty;

        var script = $@"
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

$allMailboxes = $allMailboxes | Group-Object Identity | ForEach-Object {{ $_.Group | Select-Object -First 1 }}

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
}}
";

        onLog?.Invoke("Verbose", $"Fetching deleted mailboxes (skip={request.Skip}, pageSize={request.PageSize})...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (result.Success && result.Output.Any())
        {
            var firstOutput = result.Output.First();
            var hash = firstOutput?.BaseObject as System.Collections.Hashtable;
            if (hash != null)
            {
                response.TotalCount = Convert.ToInt32(hash["TotalCount"] ?? 0);

                var mailboxes = hash["Mailboxes"] as object[];
                if (mailboxes != null)
                {
                    foreach (var mbxObj in mailboxes)
                    {
                        if (mbxObj is System.Collections.Hashtable mbxHash)
                        {
                            var item = new DeletedMailboxItemDto
                            {
                                Identity = mbxHash["Identity"]?.ToString() ?? string.Empty,
                                DisplayName = mbxHash["DisplayName"]?.ToString() ?? string.Empty,
                                UserPrincipalName = mbxHash["UserPrincipalName"]?.ToString(),
                                PrimarySmtpAddress = mbxHash["PrimarySmtpAddress"]?.ToString() ?? string.Empty,
                                RecipientTypeDetails = mbxHash["RecipientTypeDetails"]?.ToString() ?? string.Empty,
                                Alias = mbxHash["Alias"]?.ToString(),
                                DeletionType = mbxHash["DeletionType"]?.ToString() ?? string.Empty
                            };

                            response.Mailboxes.Add(item);
                        }
                    }
                }

                response.HasMore = (request.Skip + response.Mailboxes.Count) < response.TotalCount;
            }
        }

        onLog?.Invoke("Information", $"Retrieved {response.Mailboxes.Count} deleted mailboxes (total: {response.TotalCount})");

        return response;
    }

    public async Task<MailboxDetailsDto> GetMailboxDetailsAsync(
        GetMailboxDetailsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var escapedIdentity = request.Identity.Replace("'", "''");

        var script = $@"
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

$mbx = Get-Mailbox -Identity '{escapedIdentity}'

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
}}
";

        onLog?.Invoke("Verbose", $"Fetching mailbox details for {request.Identity}...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success || !result.Output.Any())
        {
            throw new InvalidOperationException($"Failed to get mailbox: {result.ErrorMessage}");
        }

        var firstOutput = result.Output.First();
        var hash = firstOutput?.BaseObject as System.Collections.Hashtable;
        if (hash == null)
        {
            throw new InvalidOperationException("Failed to parse mailbox data");
        }

        var details = new MailboxDetailsDto
        {
            Identity = hash["Identity"]?.ToString() ?? "",
            Guid = hash["Guid"]?.ToString(),
            DisplayName = hash["DisplayName"]?.ToString() ?? "",
            PrimarySmtpAddress = hash["PrimarySmtpAddress"]?.ToString() ?? "",
            UserPrincipalName = hash["UserPrincipalName"]?.ToString(),
            Alias = hash["Alias"]?.ToString(),
            RecipientType = hash["RecipientType"]?.ToString() ?? "",
            RecipientTypeDetails = hash["RecipientTypeDetails"]?.ToString() ?? "",
            EmailAddresses = ConvertToStringList(hash["EmailAddresses"]),
            RetentionPolicy = hash["RetentionPolicy"]?.ToString(),
            WhenCreated = hash["WhenCreated"] as DateTime?,
            WhenMailboxCreated = hash["WhenMailboxCreated"] as DateTime?,
            Features = new MailboxFeaturesDto
            {
                ArchiveEnabled = hash["ArchiveEnabled"] as bool? ?? false,
                ArchiveName = hash["ArchiveName"]?.ToString(),
                ArchiveGuid = hash["ArchiveGuid"]?.ToString(),
                ArchiveStatus = hash["ArchiveStatus"]?.ToString(),

                LitigationHoldEnabled = hash["LitigationHoldEnabled"] as bool? ?? false,
                LitigationHoldDate = hash["LitigationHoldDate"] as DateTime?,
                LitigationHoldOwner = hash["LitigationHoldOwner"]?.ToString(),
                LitigationHoldDuration = hash["LitigationHoldDuration"]?.ToString(),

                AuditEnabled = hash["AuditEnabled"] as bool? ?? false,
                AuditLogAgeLimit = hash["AuditLogAgeLimit"]?.ToString(),
                AuditAdmin = ConvertToStringList(hash["AuditAdmin"]),
                AuditDelegate = ConvertToStringList(hash["AuditDelegate"]),
                AuditOwner = ConvertToStringList(hash["AuditOwner"]),

                ForwardingAddress = hash["ForwardingAddress"]?.ToString(),
                ForwardingSmtpAddress = hash["ForwardingSmtpAddress"]?.ToString(),
                DeliverToMailboxAndForward = hash["DeliverToMailboxAndForward"] as bool? ?? false,

                ProhibitSendQuota = hash["ProhibitSendQuota"]?.ToString(),
                ProhibitSendQuotaBytes = hash["ProhibitSendQuotaBytes"] as long?,
                ProhibitSendReceiveQuota = hash["ProhibitSendReceiveQuota"]?.ToString(),
                ProhibitSendReceiveQuotaBytes = hash["ProhibitSendReceiveQuotaBytes"] as long?,
                IssueWarningQuota = hash["IssueWarningQuota"]?.ToString(),
                IssueWarningQuotaBytes = hash["IssueWarningQuotaBytes"] as long?,
                MaxSendSize = hash["MaxSendSize"]?.ToString(),
                MaxReceiveSize = hash["MaxReceiveSize"]?.ToString(),

                RetentionHoldEnabled = hash["RetentionHoldEnabled"] as bool? ?? false,
                SingleItemRecoveryEnabled = hash["SingleItemRecoveryEnabled"] as bool? ?? false,
                RetainDeletedItemsFor = hash["RetainDeletedItemsFor"]?.ToString(),
                HiddenFromAddressListsEnabled = hash["HiddenFromAddressListsEnabled"] as bool? ?? false
            }
        };

        cancellationToken.ThrowIfCancellationRequested();

        if (request.IncludeStatistics)
        {
            details.Statistics = await GetMailboxStatisticsAsync(request.Identity, onLog, cancellationToken);
        }

        cancellationToken.ThrowIfCancellationRequested();

        if (request.IncludeRules)
        {
            details.InboxRules = await GetInboxRulesAsync(request.Identity, onLog, cancellationToken);
        }

        cancellationToken.ThrowIfCancellationRequested();

        if (request.IncludeAutoReply)
        {
            details.AutoReplyConfiguration = await GetAutoReplyConfigurationAsync(request.Identity, onLog, cancellationToken);
        }

        cancellationToken.ThrowIfCancellationRequested();

        if (request.IncludePermissions)
        {
            details.Permissions = await GetMailboxPermissionsAsync(request.Identity, onLog, cancellationToken);
        }

        cancellationToken.ThrowIfCancellationRequested();

        // Fetch assigned licenses via Microsoft Graph (if UPN is available)
        if (!string.IsNullOrEmpty(details.UserPrincipalName))
        {
            try
            {
                onLog?.Invoke("Verbose", $"Fetching user licenses for {details.UserPrincipalName}...");
                var licenseResponse = await GetUserLicensesAsync(details.UserPrincipalName, cancellationToken);
                details.AssignedLicenses = licenseResponse.Licenses;
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                onLog?.Invoke("Warning", $"Could not fetch user licenses: {ex.Message}");
            }
        }

        onLog?.Invoke("Information", $"Retrieved details for mailbox {details.DisplayName}");

        return details;
    }

    public async Task<List<RetentionPolicySummaryDto>> GetRetentionPoliciesAsync(
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var script = @"
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
})
";

        onLog?.Invoke("Verbose", "Fetching retention policies...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to get retention policies: {result.ErrorMessage}");
        }

        var policies = new List<RetentionPolicySummaryDto>();

        foreach (var output in result.Output)
        {
            if (output.BaseObject is object[] array)
            {
                foreach (var item in array)
                {
                    if (item is System.Collections.Hashtable itemHash)
                    {
                        policies.Add(new RetentionPolicySummaryDto
                        {
                            Id = itemHash["Id"]?.ToString(),
                            Name = itemHash["Name"]?.ToString() ?? string.Empty,
                            Description = itemHash["Description"]?.ToString(),
                            RequiresArchive = itemHash["RequiresArchive"] as bool? ?? false
                        });
                    }
                }

                continue;
            }

            if (output.BaseObject is System.Collections.Hashtable hash)
            {
                policies.Add(new RetentionPolicySummaryDto
                {
                    Id = hash["Id"]?.ToString(),
                    Name = hash["Name"]?.ToString() ?? string.Empty,
                    Description = hash["Description"]?.ToString(),
                    RequiresArchive = hash["RequiresArchive"] as bool? ?? false
                });
            }
        }

        return policies;
    }

    private async Task<MailboxStatisticsDto?> GetMailboxStatisticsAsync(
        string identity,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var escapedIdentity = identity.Replace("'", "''");

        var script = $@"
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
    $stats = Get-MailboxStatistics -Identity '{escapedIdentity}' -ErrorAction Stop
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
}}
";

        onLog?.Invoke("Verbose", "Fetching mailbox statistics...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        var firstOutput = result.Output.Any() ? result.Output.First() : null;
        if (result.Success && firstOutput != null && firstOutput.BaseObject is System.Collections.Hashtable hash)
        {
            return new MailboxStatisticsDto
            {
                TotalItemSize = hash["TotalItemSize"]?.ToString(),
                TotalItemSizeBytes = hash["TotalItemSizeBytes"] as long?,
                ItemCount = Convert.ToInt32(hash["ItemCount"] ?? 0),
                DeletedItemCount = Convert.ToInt32(hash["DeletedItemCount"] ?? 0),
                TotalDeletedItemSize = hash["TotalDeletedItemSize"]?.ToString(),
                LastLogonTime = hash["LastLogonTime"] as DateTime?,
                LastLogoffTime = hash["LastLogoffTime"] as DateTime?
            };
        }

        return null;
    }

    private async Task<List<InboxRuleDto>?> GetInboxRulesAsync(
        string identity,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var escapedIdentity = identity.Replace("'", "''");

        var script = $@"
try {{
    $rules = Get-InboxRule -Mailbox '{escapedIdentity}' -ErrorAction Stop
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
}}
";

        onLog?.Invoke("Verbose", "Fetching inbox rules...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.Success && result.Output.Any())
        {
            var rules = new List<InboxRuleDto>();

            foreach (var output in result.Output)
            {
                if (output.BaseObject is System.Collections.Hashtable hash)
                {
                    rules.Add(new InboxRuleDto
                    {
                        Name = hash["Name"]?.ToString() ?? "",
                        RuleIdentity = hash["RuleIdentity"]?.ToString(),
                        Enabled = hash["Enabled"] as bool? ?? false,
                        Priority = Convert.ToInt32(hash["Priority"] ?? 0),
                        Description = hash["Description"]?.ToString(),
                        ForwardTo = ConvertToStringList(hash["ForwardTo"]),
                        ForwardAsAttachmentTo = ConvertToStringList(hash["ForwardAsAttachmentTo"]),
                        RedirectTo = ConvertToStringList(hash["RedirectTo"]),
                        DeleteMessage = hash["DeleteMessage"] as bool? ?? false,
                        MoveToFolder = hash["MoveToFolder"]?.ToString()
                    });
                }
            }

            return rules;
        }

        return null;
    }

    private async Task<AutoReplyConfigurationDto?> GetAutoReplyConfigurationAsync(
        string identity,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var escapedIdentity = identity.Replace("'", "''");

        var script = $@"
try {{
    $config = Get-MailboxAutoReplyConfiguration -Identity '{escapedIdentity}' -ErrorAction Stop
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
}}
";

        onLog?.Invoke("Verbose", "Fetching auto-reply configuration...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        var firstOutput = result.Output.Any() ? result.Output.First() : null;
        if (result.Success && firstOutput != null && firstOutput.BaseObject is System.Collections.Hashtable hash)
        {
            return new AutoReplyConfigurationDto
            {
                AutoReplyState = hash["AutoReplyState"]?.ToString() ?? "Disabled",
                StartTime = hash["StartTime"] as DateTime?,
                EndTime = hash["EndTime"] as DateTime?,
                InternalMessage = hash["InternalMessage"]?.ToString(),
                ExternalMessage = hash["ExternalMessage"]?.ToString(),
                ExternalAudience = hash["ExternalAudience"]?.ToString()
            };
        }

        return null;
    }

    public async Task SetMailboxSettingsAsync(
        UpdateMailboxSettingsRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var escapedIdentity = request.Identity.Replace("'", "''");
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

        if (scriptBuilder.Length == 0)
        {
            return;
        }

        onLog?.Invoke("Information", $"Updating mailbox settings for {request.Identity}...");

        var result = await _engine.ExecuteAsync(scriptBuilder.ToString(), onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to update mailbox settings: {result.ErrorMessage}");
        }

        onLog?.Invoke("Information", "Mailbox settings updated successfully");
    }

    public async Task SetRetentionPolicyAsync(
        SetRetentionPolicyRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var escapedIdentity = request.Identity.Replace("'", "''");
        var policyValue = FormatNullableString(request.PolicyName);
        var script = $"Set-Mailbox -Identity '{escapedIdentity}' -RetentionPolicy {policyValue}";

        onLog?.Invoke("Information", $"Updating retention policy for {request.Identity}...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to update retention policy: {result.ErrorMessage}");
        }

        onLog?.Invoke("Information", "Retention policy updated successfully");
    }

    public async Task SetMailboxAutoReplyConfigurationAsync(
        SetMailboxAutoReplyConfigurationRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var escapedIdentity = request.Identity.Replace("'", "''");
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
            scriptBuilder.AppendLine($"$params.ExternalAudience = '{request.ExternalAudience.Replace("'", "''")}'");
        }

        scriptBuilder.AppendLine("Set-MailboxAutoReplyConfiguration @params");

        onLog?.Invoke("Information", $"Updating auto-reply configuration for {request.Identity}...");

        var result = await _engine.ExecuteAsync(scriptBuilder.ToString(), onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to update auto-reply configuration: {result.ErrorMessage}");
        }

        onLog?.Invoke("Information", "Auto-reply configuration updated successfully");
    }

    public async Task ConvertMailboxToSharedAsync(
        ConvertMailboxToSharedRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var escapedIdentity = request.Identity.Replace("'", "''");
        var script = $"Set-Mailbox -Identity '{escapedIdentity}' -Type Shared";

        onLog?.Invoke("Information", $"Converting mailbox {request.Identity} to shared...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to convert mailbox: {result.ErrorMessage}");
        }

        onLog?.Invoke("Information", "Mailbox converted to shared successfully");
    }

    public async Task ConvertMailboxToRegularAsync(
        ConvertMailboxToRegularRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var escapedIdentity = request.Identity.Replace("'", "''");
        var script = $"Set-Mailbox -Identity '{escapedIdentity}' -Type Regular";

        onLog?.Invoke("Information", $"Converting mailbox {request.Identity} to regular...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to convert mailbox: {result.ErrorMessage}");
        }

        onLog?.Invoke("Information", "Mailbox converted to regular successfully");
    }

    public async Task<RestoreMailboxResponse> RestoreMailboxAsync(
        RestoreMailboxRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
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
                $result.Scenario = 'HardDeleted'
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

    $result.Success = $true
    return $result
}
catch {
    $message = $_.Exception.Message
    $code = Resolve-ErrorCode $message
    return Set-ErrorResult $code $message
}
";

        onLog?.Invoke("Information", $"Starting mailbox restore for {request.SourceIdentity}...");

        var parameters = new Dictionary<string, object>
        {
            ["SourceIdentity"] = request.SourceIdentity,
            ["TargetMailbox"] = request.TargetMailbox ?? string.Empty,
            ["AllowLegacyDnMismatch"] = request.AllowLegacyDnMismatch
        };

        var result = await _engine.ExecuteAsync(
            script,
            parameters,
            onVerbose: onLog,
            cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success && result.Output.Count == 0)
        {
            throw new InvalidOperationException($"Failed to restore mailbox: {result.ErrorMessage}");
        }

        static RestoreMailboxStatus MapRestoreStatus(string value)
        {
            if (Enum.TryParse(value, true, out RestoreMailboxStatus parsed))
            {
                return parsed;
            }

            return value.ToLowerInvariant() switch
            {
                "queued" => RestoreMailboxStatus.InProgress,
                "inprogress" => RestoreMailboxStatus.InProgress,
                "completed" => RestoreMailboxStatus.Completed,
                "failed" => RestoreMailboxStatus.Failed,
                "suspended" => RestoreMailboxStatus.Failed,
                _ => RestoreMailboxStatus.NotStarted
            };
        }

        var response = new RestoreMailboxResponse
        {
            SourceIdentity = request.SourceIdentity,
            TargetMailbox = request.TargetMailbox,
            Scenario = RestoreMailboxScenario.Unknown,
            Status = RestoreMailboxStatus.NotStarted
        };

        if (result.Output.Count > 0 && result.Output.First().BaseObject is System.Collections.Hashtable hash)
        {
            var scenarioValue = hash["Scenario"]?.ToString();
            if (!string.IsNullOrWhiteSpace(scenarioValue) &&
                Enum.TryParse(scenarioValue, true, out RestoreMailboxScenario parsedScenario))
            {
                response.Scenario = parsedScenario;
            }

            response.Action = hash["Action"]?.ToString();

            var statusValue = hash["Status"]?.ToString();
            if (!string.IsNullOrWhiteSpace(statusValue))
            {
                response.Status = MapRestoreStatus(statusValue);
            }

            response.StatusDetail = hash["StatusDetail"]?.ToString();

            if (hash["PercentComplete"] != null &&
                int.TryParse(hash["PercentComplete"]?.ToString(), out var percent))
            {
                response.PercentComplete = percent;
            }

            response.RequestGuid = hash["RequestGuid"]?.ToString();

            var errorCodeValue = hash["ErrorCode"]?.ToString();
            if (!string.IsNullOrWhiteSpace(errorCodeValue) &&
                Enum.TryParse(errorCodeValue, true, out RestoreMailboxErrorCode parsedErrorCode) &&
                parsedErrorCode != RestoreMailboxErrorCode.None)
            {
                response.Error = new RestoreMailboxErrorDto
                {
                    Code = parsedErrorCode,
                    Message = hash["ErrorMessage"]?.ToString() ?? "Mailbox restore failed"
                };
            }

            if (response.Error == null && response.Status == RestoreMailboxStatus.Failed)
            {
                response.Error = new RestoreMailboxErrorDto
                {
                    Code = RestoreMailboxErrorCode.Unknown,
                    Message = hash["ErrorMessage"]?.ToString() ?? "Mailbox restore failed"
                };
            }

            response.SourceIdentity = hash["SourceIdentity"]?.ToString() ?? response.SourceIdentity;
            response.TargetMailbox = hash["TargetMailbox"]?.ToString() ?? response.TargetMailbox;
        }
        else if (!result.Success)
        {
            response.Status = RestoreMailboxStatus.Failed;
            response.Error = new RestoreMailboxErrorDto
            {
                Code = RestoreMailboxErrorCode.Unknown,
                Message = result.ErrorMessage ?? "Mailbox restore failed"
            };
        }

        return response;
    }

    public async Task<GetMailboxSpaceReportResponse> GetMailboxSpaceReportAsync(
        GetMailboxSpaceReportRequest request,
        Action<string, string>? onLog,
        Action<int, int>? onProgress,
        CancellationToken cancellationToken)
    {
        var script = @"
$ErrorActionPreference = 'SilentlyContinue'

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

$mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox -ErrorAction SilentlyContinue
$total = @($mailboxes).Count
$index = 0

foreach ($mbx in $mailboxes) {
    $index++
    $stats = $null
    try {
        $stats = Get-MailboxStatistics -Identity $mbx.Identity -ErrorAction SilentlyContinue
    } catch {
        $stats = $null
    }

    @{
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
";

        onLog?.Invoke("Information", "Generating mailbox space report...");

        var result = await _engine.ExecuteAsync(
            script,
            onVerbose: onLog,
            onOutput: output =>
            {
                if (onProgress == null)
                {
                    return;
                }

                if (output?.BaseObject is System.Collections.Hashtable hash &&
                    hash["Index"] != null &&
                    hash["TotalCount"] != null)
                {
                    var index = Convert.ToInt32(hash["Index"]);
                    var total = Convert.ToInt32(hash["TotalCount"]);
                    onProgress(index, total);
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

        var response = new GetMailboxSpaceReportResponse();

        foreach (var output in result.Output)
        {
            if (output?.BaseObject is System.Collections.Hashtable hash)
            {
                response.Mailboxes.Add(new MailboxSpaceItemDto
                {
                    Identity = hash["Identity"]?.ToString() ?? string.Empty,
                    DisplayName = hash["DisplayName"]?.ToString() ?? string.Empty,
                    PrimarySmtpAddress = hash["PrimarySmtpAddress"]?.ToString() ?? string.Empty,
                    TotalItemSize = hash["TotalItemSize"]?.ToString(),
                    TotalItemSizeBytes = hash["TotalItemSizeBytes"] as long?,
                    ProhibitSendQuota = hash["ProhibitSendQuota"]?.ToString(),
                    ProhibitSendQuotaBytes = hash["ProhibitSendQuotaBytes"] as long?,
                    ProhibitSendReceiveQuota = hash["ProhibitSendReceiveQuota"]?.ToString(),
                    ProhibitSendReceiveQuotaBytes = hash["ProhibitSendReceiveQuotaBytes"] as long?,
                    IssueWarningQuota = hash["IssueWarningQuota"]?.ToString(),
                    IssueWarningQuotaBytes = hash["IssueWarningQuotaBytes"] as long?
                });
            }
        }

        onLog?.Invoke("Information", $"Mailbox space report generated: {response.Mailboxes.Count} entries");

        return response;
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

        onLog?.Invoke("Verbose", "Fetching FullAccess permissions...");

        var fullAccessResult = await _engine.ExecuteAsync(fullAccessScript, onVerbose: onLog, cancellationToken: cancellationToken);

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
            if ($recipient.DisplayName) {{
                $displayName = $recipient.DisplayName
            }}
            if ($recipient.PrimarySmtpAddress) {{
                $trusteeIdentity = $recipient.PrimarySmtpAddress.ToString()
            }}
            elseif ($recipient.ExternalDirectoryObjectId) {{
                $trusteeIdentity = $recipient.ExternalDirectoryObjectId.ToString()
            }}
            elseif ($recipient.Identity) {{
                $trusteeIdentity = $recipient.Identity.ToString()
            }}
        }}
        catch {{
        }}
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

        onLog?.Invoke("Verbose", "Fetching SendAs permissions...");

        var sendAsResult = await _engine.ExecuteAsync(sendAsScript, onVerbose: onLog, cancellationToken: cancellationToken);

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
            if ($recipient.DisplayName) {{
                $displayName = $recipient.DisplayName
            }}
            elseif ($recipient.PrimarySmtpAddress) {{
                $displayName = $recipient.PrimarySmtpAddress.ToString()
            }}
        }}
        catch {{
        }}
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

        onLog?.Invoke("Verbose", "Fetching SendOnBehalf permissions...");

        var sendOnBehalfResult = await _engine.ExecuteAsync(sendOnBehalfScript, onVerbose: onLog, cancellationToken: cancellationToken);

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

        onLog?.Invoke("Information", $"Retrieved permissions: {permissions.FullAccessPermissions.Count} FullAccess, {permissions.SendAsPermissions.Count} SendAs, {permissions.SendOnBehalfPermissions.Count} SendOnBehalf");

        return permissions;
    }

    public async Task SetMailboxPermissionAsync(
        SetMailboxPermissionRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var escapedIdentity = request.Identity.Replace("'", "''");
        var escapedUser = request.User.Replace("'", "''");

        string script;

        switch (request.PermissionType)
        {
            case PermissionType.FullAccess:
                if (request.Action == PermissionAction.Add)
                {
                    var autoMapping = request.AutoMapping ?? true;
                    script = $"Add-MailboxPermission -Identity '{escapedIdentity}' -User '{escapedUser}' -AccessRights FullAccess -AutoMapping ${autoMapping} -Confirm:$false";
                }
                else if (request.Action == PermissionAction.Modify)
                {
                    var autoMapping = request.AutoMapping ?? true;
                    script = $@"
Remove-MailboxPermission -Identity '{escapedIdentity}' -User '{escapedUser}' -AccessRights FullAccess -Confirm:$false
Add-MailboxPermission -Identity '{escapedIdentity}' -User '{escapedUser}' -AccessRights FullAccess -AutoMapping ${autoMapping} -Confirm:$false
";
                }
                else
                {
                    script = $"Remove-MailboxPermission -Identity '{escapedIdentity}' -User '{escapedUser}' -AccessRights FullAccess -Confirm:$false";
                }
                break;

            case PermissionType.SendAs:
                if (request.Action == PermissionAction.Add)
                {
                    script = $"Add-RecipientPermission -Identity '{escapedIdentity}' -Trustee '{escapedUser}' -AccessRights SendAs -Confirm:$false -ErrorAction Stop";
                }
                else
                {
                    script = $"Remove-RecipientPermission -Identity '{escapedIdentity}' -Trustee '{escapedUser}' -AccessRights SendAs -Confirm:$false -ErrorAction Stop";
                }
                break;

            case PermissionType.SendOnBehalf:
                if (request.Action == PermissionAction.Add)
                {
                    script = $@"
$mbx = Get-Mailbox -Identity '{escapedIdentity}'
$current = @($mbx.GrantSendOnBehalfTo)
$current += '{escapedUser}'
Set-Mailbox -Identity '{escapedIdentity}' -GrantSendOnBehalfTo $current
";
                }
                else
                {
                    script = $@"
$mbx = Get-Mailbox -Identity '{escapedIdentity}'
$current = @($mbx.GrantSendOnBehalfTo) | Where-Object {{ $_.ToString() -ne '{escapedUser}' }}
Set-Mailbox -Identity '{escapedIdentity}' -GrantSendOnBehalfTo $current
";
                }
                break;

            default:
                throw new ArgumentException($"Unknown permission type: {request.PermissionType}");
        }

        var actionVerb = request.Action switch
        {
            PermissionAction.Add => "Adding",
            PermissionAction.Remove => "Removing",
            PermissionAction.Modify => "Modifying",
            _ => "Processing"
        };
        onLog?.Invoke("Information", $"{actionVerb} {request.PermissionType} permission for {request.User} on {request.Identity}...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to {request.Action} {request.PermissionType} permission: {result.ErrorMessage}");
        }

        var actionPastTense = request.Action switch
        {
            PermissionAction.Add => "added",
            PermissionAction.Remove => "removed",
            PermissionAction.Modify => "modified",
            _ => "processed"
        };
        onLog?.Invoke("Information", $"Successfully {actionPastTense} {request.PermissionType} permission");
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

        onLog?.Invoke("Information", $"Applying permissions delta plan: {request.Actions.Count} actions");

        for (int i = 0; i < request.Actions.Count; i++)
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
                onLog?.Invoke("Error", $"Failed: {action.Description} - {ex.Message}");
            }

            response.Results.Add(result);
            onProgress?.Invoke(i + 1, request.Actions.Count);
        }

        onLog?.Invoke("Information", $"Delta plan complete: {response.SuccessfulActions} succeeded, {response.FailedActions} failed");

        return response;
    }

    public async Task<List<TenantLicenseDto>> GetTenantLicensesAsync(CancellationToken cancellationToken = default)
    {
        var script = @"
try {
    $skus = Get-MgSubscribedSku -ErrorAction Stop
    foreach ($sku in $skus) {
        [PSCustomObject]@{
            SkuId = $sku.SkuId
            SkuPartNumber = $sku.SkuPartNumber
            Total = $sku.PrepaidUnits.Enabled
            Assigned = $sku.ConsumedUnits
            Available = ($sku.PrepaidUnits.Enabled - $sku.ConsumedUnits)
        }
    }
} catch {
    Write-Warning ""Get-MgSubscribedSku not available: $($_.Exception.Message)""
}";
        var results = await RunScriptAllowErrorsAsync(script, cancellationToken);
        var licenses = new List<TenantLicenseDto>();
        foreach (var obj in results)
        {
            licenses.Add(new TenantLicenseDto
            {
                SkuId = GetString(obj, "SkuId"),
                SkuPartNumber = GetString(obj, "SkuPartNumber"),
                DisplayName = GetString(obj, "SkuPartNumber"),
                Total = GetInt(obj, "Total"),
                Assigned = GetInt(obj, "Assigned"),
                Available = GetInt(obj, "Available")
            });
        }
        return licenses;
    }

    public async Task<List<AdminRoleMemberDto>> GetAdminRoleMembersAsync(CancellationToken cancellationToken = default)
    {
        var script = @"
try {
    $adminRoles = @('Global Administrator', 'Exchange Administrator', 'User Administrator', 'Security Administrator', 'Helpdesk Administrator', 'SharePoint Administrator', 'Teams Administrator', 'Billing Administrator')
    $results = @()
    foreach ($roleName in $adminRoles) {
        try {
            $role = Get-MgDirectoryRole -Filter ""displayName eq '$roleName'"" -ErrorAction SilentlyContinue
            if ($role) {
                $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -ErrorAction SilentlyContinue
                foreach ($member in $members) {
                    $results += [PSCustomObject]@{
                        DisplayName = $member.AdditionalProperties['displayName']
                        UserPrincipalName = $member.AdditionalProperties['userPrincipalName']
                        RoleName = $roleName
                    }
                }
            }
        } catch {
            continue
        }
    }
    $results
} catch {
    Write-Warning ""Get-MgDirectoryRole not available: $($_.Exception.Message)""
}";
        var results = await RunScriptAllowErrorsAsync(script, cancellationToken);
        var admins = new List<AdminRoleMemberDto>();
        foreach (var obj in results)
        {
            admins.Add(new AdminRoleMemberDto
            {
                DisplayName = GetString(obj, "DisplayName"),
                UserPrincipalName = GetString(obj, "UserPrincipalName"),
                RoleName = GetString(obj, "RoleName")
            });
        }
        return admins;
    }

    public async Task<GetMessageTraceResponse> GetMessageTraceAsync(GetMessageTraceRequest request, CancellationToken cancellationToken = default)
    {
        var script = $@"
$WarningPreference = 'SilentlyContinue'
$pageSize = {request.PageSize}
$page = {request.Page}
$take = $pageSize + 1
$params = @{{
    ResultSize = ($pageSize * $page) + 1
    StartDate = [DateTime]::Parse('{request.StartDate:o}')
    EndDate = [DateTime]::Parse('{request.EndDate:o}')
}}
if ('{request.SenderAddress ?? ""}' -ne '') {{ $params['SenderAddress'] = '{request.SenderAddress}' }}
if ('{request.RecipientAddress ?? ""}' -ne '') {{ $params['RecipientAddress'] = '{request.RecipientAddress}' }}

# Use Get-MessageTraceV2 (Get-MessageTrace is deprecated)
if (-not (Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue)) {{
    throw 'Get-MessageTraceV2 is not available. Install/upgrade ExchangeOnlineManagement and reconnect.'
}}

$traces = Get-MessageTraceV2 @params -ErrorAction Stop |
    Select-Object -Skip (($page - 1) * $pageSize) -First $take

foreach ($t in $traces) {{
    [PSCustomObject]@{{
        MessageId = $t.MessageId
        MessageTraceId = $t.MessageTraceId.ToString()
        SenderAddress = $t.SenderAddress
        RecipientAddress = $t.RecipientAddress
        Subject = $t.Subject
        Status = $t.Status
        Received = $t.Received.ToString('o')
        Size = $t.Size
    }}
}}";
        var results = await RunScriptAsync(script, cancellationToken);
        var messages = new List<MessageTraceItemDto>();
        foreach (var obj in results)
        {
            messages.Add(new MessageTraceItemDto
            {
                MessageId = GetString(obj, "MessageId"),
                MessageTraceId = GetString(obj, "MessageTraceId"),
                SenderAddress = GetString(obj, "SenderAddress"),
                RecipientAddress = GetString(obj, "RecipientAddress"),
                Subject = GetString(obj, "Subject"),
                Status = GetString(obj, "Status"),
                Received = GetNullableDateTime(obj, "Received"),
                Size = GetNullableLong(obj, "Size")
            });
        }
        var hasMore = false;
        if (messages.Count > request.PageSize)
        {
            hasMore = true;
            messages.RemoveAt(messages.Count - 1);
        }
        return new GetMessageTraceResponse
        {
            Messages = messages,
            TotalCount = ((request.Page - 1) * request.PageSize) + messages.Count,
            HasMore = hasMore,
            Page = request.Page,
            PageSize = request.PageSize
        };
    }

    public async Task<GetUserLicensesResponse> GetUserLicensesAsync(string userPrincipalName, CancellationToken cancellationToken = default)
    {
        var script = $@"
try {{
    $licenses = Get-MgUserLicenseDetail -UserId '{userPrincipalName}' -ErrorAction Stop
    foreach ($lic in $licenses) {{
        [PSCustomObject]@{{
            SkuId = $lic.SkuId
            SkuPartNumber = $lic.SkuPartNumber
        }}
    }}
}} catch {{
    Write-Warning ""Get-MgUserLicenseDetail not available: $($_.Exception.Message)""
}}";
        var results = await RunScriptAsync(script, cancellationToken);
        var licenses = new List<UserLicenseDto>();
        foreach (var obj in results)
        {
            licenses.Add(new UserLicenseDto
            {
                SkuId = GetString(obj, "SkuId"),
                SkuPartNumber = GetString(obj, "SkuPartNumber"),
                DisplayName = GetString(obj, "SkuPartNumber")
            });
        }
        return new GetUserLicensesResponse { Licenses = licenses };
    }

    public async Task SetUserLicenseAsync(SetUserLicenseRequest request, CancellationToken cancellationToken = default)
    {
        var addSkus = string.Join(",", request.AddLicenseSkuIds.Select(s => $"@{{SkuId='{s}'}}"));
        var removeSkus = string.Join(",", request.RemoveLicenseSkuIds.Select(s => $"'{s}'"));
        var script = $@"
$addLicenses = @({addSkus})
$removeLicenses = @({removeSkus})
Set-MgUserLicense -UserId '{request.UserPrincipalName}' -AddLicenses $addLicenses -RemoveLicenses $removeLicenses -ErrorAction Stop
Write-Output 'License updated successfully'";
        await RunScriptAsync(script, cancellationToken);
    }

    public async Task<GetAvailableLicensesResponse> GetAvailableLicensesAsync(CancellationToken cancellationToken = default)
    {
        var script = @"
try {
    Get-MgSubscribedSku -ErrorAction Stop | Where-Object { ($_.PrepaidUnits.Enabled - $_.ConsumedUnits) -gt 0 } | ForEach-Object {
        [PSCustomObject]@{
            SkuId = $_.SkuId
            SkuPartNumber = $_.SkuPartNumber
            Total = $_.PrepaidUnits.Enabled
            Assigned = $_.ConsumedUnits
            Available = ($_.PrepaidUnits.Enabled - $_.ConsumedUnits)
        }
    }
} catch {
    Write-Warning ""Get-MgSubscribedSku not available: $($_.Exception.Message)""
}";
        var results = await RunScriptAsync(script, cancellationToken);
        var licenses = new List<TenantLicenseDto>();
        foreach (var obj in results)
        {
            licenses.Add(new TenantLicenseDto
            {
                SkuId = GetString(obj, "SkuId"),
                SkuPartNumber = GetString(obj, "SkuPartNumber"),
                DisplayName = GetString(obj, "SkuPartNumber"),
                Total = GetInt(obj, "Total"),
                Assigned = GetInt(obj, "Assigned"),
                Available = GetInt(obj, "Available")
            });
        }
        return new GetAvailableLicensesResponse { Licenses = licenses };
    }

    public async Task<PrerequisiteStatusDto> CheckPrerequisitesAsync(CancellationToken cancellationToken = default)
    {
        var script = @"
$psVersion = $PSVersionTable.PSVersion.ToString()
$isPwsh7 = $PSVersionTable.PSVersion.Major -ge 7

$exoModule = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Select-Object -First 1
$graphModule = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Select-Object -First 1

[PSCustomObject]@{
    PowerShellVersion = $psVersion
    IsPowerShell7 = $isPwsh7
    ExchangeModuleInstalled = ($null -ne $exoModule)
    ExchangeModuleVersion = if ($exoModule) { $exoModule.Version.ToString() } else { $null }
    GraphModuleInstalled = ($null -ne $graphModule)
    GraphModuleVersion = if ($graphModule) { $graphModule.Version.ToString() } else { $null }
}";
        var results = await RunScriptAsync(script, cancellationToken);
        if (results.Count == 0)
        {
            return new PrerequisiteStatusDto();
        }
        var obj = results[0];
        return new PrerequisiteStatusDto
        {
            PowerShellVersion = GetString(obj, "PowerShellVersion"),
            IsPowerShell7 = GetBool(obj, "IsPowerShell7"),
            ExchangeModuleInstalled = GetBool(obj, "ExchangeModuleInstalled"),
            ExchangeModuleVersion = GetNullableString(obj, "ExchangeModuleVersion"),
            GraphModuleInstalled = GetBool(obj, "GraphModuleInstalled"),
            GraphModuleVersion = GetNullableString(obj, "GraphModuleVersion")
        };
    }

    public async Task<InstallModuleResponse> InstallModuleAsync(string moduleName, CancellationToken cancellationToken = default)
    {
        var script = $@"
try {{
    Write-Output 'Installing {moduleName}...'
    $nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
    if (-not $nuget) {{
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop
    }}
    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction SilentlyContinue
    Install-Module -Name '{moduleName}' -Force -AllowClobber -Scope CurrentUser -Confirm:$false -ErrorAction Stop
    $installed = Get-Module -ListAvailable -Name '{moduleName}' | Select-Object -First 1
    [PSCustomObject]@{{
        Success = $true
        Message = '{moduleName} installed successfully'
        InstalledVersion = if ($installed) {{ $installed.Version.ToString() }} else {{ $null }}
    }}
}} catch {{
    [PSCustomObject]@{{
        Success = $false
        Message = ""Failed to install {moduleName}: $($_.Exception.Message)""
        InstalledVersion = $null
    }}
}}";
        var result = await _engine.ExecuteAsync(script, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }
        if (result.Output.Count == 0)
        {
            return new InstallModuleResponse
            {
                Success = false,
                Message = result.ErrorMessage ?? $"No output from Install-Module {moduleName}",
                ModuleName = moduleName
            };
        }
        var obj = result.Output.Last();
        var success = GetBool(obj, "Success");
        var message = GetString(obj, "Message");
        if (!success && string.IsNullOrWhiteSpace(message))
        {
            message = result.ErrorMessage ?? $"Failed to install {moduleName}";
        }
        return new InstallModuleResponse
        {
            Success = success,
            Message = message,
            ModuleName = moduleName,
            InstalledVersion = GetNullableString(obj, "InstalledVersion")
        };
    }

    private async Task<List<PSObject>> RunScriptAsync(string script, CancellationToken cancellationToken = default)
    {
        var result = await _engine.ExecuteAsync(script, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }
        if (!result.Success)
        {
            var errorMsg = result.ErrorMessage ?? "PowerShell script execution failed";
            throw new InvalidOperationException(errorMsg);
        }
        return result.Output;
    }

    private async Task<List<PSObject>> RunScriptAllowErrorsAsync(string script, CancellationToken cancellationToken = default)
    {
        var result = await _engine.ExecuteAsync(script, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }
        return result.Output;
    }

    private static string GetString(PSObject obj, string propertyName)
    {
        return obj.Properties[propertyName]?.Value?.ToString() ?? string.Empty;
    }

    private static string? GetNullableString(PSObject obj, string propertyName)
    {
        return obj.Properties[propertyName]?.Value?.ToString();
    }

    private static int GetInt(PSObject obj, string propertyName)
    {
        var value = obj.Properties[propertyName]?.Value;
        if (value == null) return 0;
        return Convert.ToInt32(value);
    }

    private static bool GetBool(PSObject obj, string propertyName)
    {
        var value = obj.Properties[propertyName]?.Value;
        if (value == null) return false;
        if (value is bool b) return b;
        return Convert.ToBoolean(value);
    }

    private static DateTime? GetNullableDateTime(PSObject obj, string propertyName)
    {
        var value = obj.Properties[propertyName]?.Value;
        if (value == null) return null;
        if (value is DateTime dt) return dt;
        if (DateTime.TryParse(value.ToString(), out var parsed)) return parsed;
        return null;
    }

    private static long? GetNullableLong(PSObject obj, string propertyName)
    {
        var value = obj.Properties[propertyName]?.Value;
        if (value == null) return null;
        if (value is long l) return l;
        return Convert.ToInt64(value);
    }

    private static List<string> ConvertToStringList(object? obj)
    {
        if (obj == null) return new List<string>();

        if (obj is object[] array)
        {
            return array.Select(x => x?.ToString() ?? "").Where(x => !string.IsNullOrEmpty(x)).ToList();
        }

        if (obj is System.Collections.IEnumerable enumerable)
        {
            var list = new List<string>();
            foreach (var item in enumerable)
            {
                var str = item?.ToString();
                if (!string.IsNullOrEmpty(str))
                {
                    list.Add(str);
                }
            }
            return list;
        }

        var single = obj.ToString();
        return string.IsNullOrEmpty(single) ? new List<string>() : new List<string> { single };
    }

    private static string FormatNullableString(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "$null";
        }

        return $"'{value.Replace("'", "''")}'";
    }

    private static string FormatNullableMessage(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "$null";
        }

        var base64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(value));
        return $"[System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('{base64}'))";
    }
}
