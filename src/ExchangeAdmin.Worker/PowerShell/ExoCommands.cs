using System.Management.Automation;
using System.Text;
using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

/// <summary>
/// Exchange Online command execution helpers.
/// </summary>
public class ExoCommands
{
    private readonly PowerShellEngine _engine;
    private readonly CapabilityDetector _capabilityDetector;

    public ExoCommands(PowerShellEngine engine, CapabilityDetector capabilityDetector)
    {
        _engine = engine;
        _capabilityDetector = capabilityDetector;
    }

    #region Dashboard

    /// <summary>
    /// Gets dashboard statistics.
    /// </summary>
    public async Task<DashboardStatsDto> GetDashboardStatsAsync(
        GetDashboardStatsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var stats = new DashboardStatsDto();
        var warnings = new List<string>();

        onLog?.Invoke("Verbose", "Fetching mailbox counts...");

        // Get mailbox counts by type
        var mailboxScript = @"
$counts = @{
    UserMailboxes = 0
    SharedMailboxes = 0
    RoomMailboxes = 0
    EquipmentMailboxes = 0
    IsApproximate = $false
}

try {
    # Try to get counts efficiently
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
    # Fallback: get approximate count
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

        // Get distribution group counts
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

        // Get unified groups if requested and available
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

        // Check for large tenant
        var totalMailboxes = stats.MailboxCounts.Total;
        if (totalMailboxes > stats.LargeTenantThreshold)
        {
            stats.IsLargeTenant = true;
            warnings.Add($"Large tenant detected ({totalMailboxes} mailboxes). Some operations may be slower.");
        }

        stats.Warnings = warnings;
        stats.RetrievedAt = DateTime.UtcNow;

        return stats;
    }

    #endregion

    #region Mailboxes

    /// <summary>
    /// Gets mailbox list with paging.
    /// </summary>
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

        // Build filter
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

        // Build script
        var script = $@"
$allMailboxes = Get-Mailbox -ResultSize Unlimited {filterParam}
";

        // Add search if specified
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

        // Sort
        var sortProperty = string.IsNullOrWhiteSpace(request.SortBy) ? "DisplayName" : request.SortBy;
        var sortDirection = request.SortDescending ? "-Descending" : "";
        script += $@"
$allMailboxes = $allMailboxes | Sort-Object {sortProperty} {sortDirection}
$totalCount = @($allMailboxes).Count

# Apply paging
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

    /// <summary>
    /// Gets mailbox details.
    /// </summary>
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
    if ($text -match '\(([\d,]+) bytes\)') {{
        return [long]($Matches[1] -replace ',', '')
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
    WhenCreated = $mbx.WhenCreated
    WhenMailboxCreated = $mbx.WhenMailboxCreated

    # Features
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

        // Get statistics if requested
        if (request.IncludeStatistics)
        {
            details.Statistics = await GetMailboxStatisticsAsync(request.Identity, onLog, cancellationToken);
        }

        cancellationToken.ThrowIfCancellationRequested();

        // Get inbox rules if requested
        if (request.IncludeRules)
        {
            details.InboxRules = await GetInboxRulesAsync(request.Identity, onLog, cancellationToken);
        }

        cancellationToken.ThrowIfCancellationRequested();

        // Get auto-reply if requested
        if (request.IncludeAutoReply)
        {
            details.AutoReplyConfiguration = await GetAutoReplyConfigurationAsync(request.Identity, onLog, cancellationToken);
        }

        cancellationToken.ThrowIfCancellationRequested();

        // Get permissions if requested
        if (request.IncludePermissions)
        {
            details.Permissions = await GetMailboxPermissionsAsync(request.Identity, onLog, cancellationToken);
        }

        onLog?.Invoke("Information", $"Retrieved details for mailbox {details.DisplayName}");

        return details;
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
    if ($text -match '\(([\d,]+) bytes\)') {{
        return [long]($Matches[1] -replace ',', '')
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

    #endregion

    #region Mailbox Updates

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

    public async Task<GetMailboxSpaceReportResponse> GetMailboxSpaceReportAsync(
        GetMailboxSpaceReportRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var script = @"
function Get-BytesFromSize($size) {
    if ($null -eq $size) { return $null }
    $text = $size.ToString()
    if ($text -match '\(([\d,]+) bytes\)') {
        return [long]($Matches[1] -replace ',', '')
    }
    return $null
}

$mailboxes = Get-Mailbox -ResultSize Unlimited
$report = @()

foreach ($mbx in $mailboxes) {
    $stats = $null
    try {
        $stats = Get-MailboxStatistics -Identity $mbx.Identity -ErrorAction Stop
    } catch {
        $stats = $null
    }

    $report += @{
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

$report
";

        onLog?.Invoke("Information", "Generating mailbox space report...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

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

    #endregion

    #region Permissions

    /// <summary>
    /// Gets mailbox permissions.
    /// </summary>
    public async Task<MailboxPermissionsDto> GetMailboxPermissionsAsync(
        string identity,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var permissions = new MailboxPermissionsDto();
        var escapedIdentity = identity.Replace("'", "''");

        // Get FullAccess permissions
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

        // Get SendAs permissions
        var sendAsScript = $@"
try {{
    $perms = Get-RecipientPermission -Identity '{escapedIdentity}' -ErrorAction Stop |
        Where-Object {{ $_.Trustee -notlike 'NT AUTHORITY\*' -and $_.Trustee -notlike 'S-1-5-*' }}

    @($perms | ForEach-Object {{
        @{{
            Identity = $_.Identity.ToString()
            Trustee = $_.Trustee.ToString()
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
                        AccessControlType = hash["AccessControlType"]?.ToString() ?? "",
                        AccessRights = ConvertToStringList(hash["AccessRights"]),
                        IsInherited = hash["IsInherited"] as bool? ?? false
                    });
                }
            }
        }

        cancellationToken.ThrowIfCancellationRequested();

        // Get SendOnBehalf from mailbox
        var sendOnBehalfScript = $@"
try {{
    $mbx = Get-Mailbox -Identity '{escapedIdentity}' -ErrorAction Stop
    @($mbx.GrantSendOnBehalfTo | ForEach-Object {{ $_.ToString() }})
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
                var value = output.BaseObject?.ToString();
                if (!string.IsNullOrEmpty(value))
                {
                    permissions.SendOnBehalfPermissions.Add(value);
                }
            }
        }

        onLog?.Invoke("Information", $"Retrieved permissions: {permissions.FullAccessPermissions.Count} FullAccess, {permissions.SendAsPermissions.Count} SendAs, {permissions.SendOnBehalfPermissions.Count} SendOnBehalf");

        return permissions;
    }

    /// <summary>
    /// Sets a mailbox permission.
    /// </summary>
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
                    script = $"Add-RecipientPermission -Identity '{escapedIdentity}' -Trustee '{escapedUser}' -AccessRights SendAs -Confirm:$false";
                }
                else
                {
                    script = $"Remove-RecipientPermission -Identity '{escapedIdentity}' -Trustee '{escapedUser}' -AccessRights SendAs -Confirm:$false";
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

    /// <summary>
    /// Applies a permissions delta plan.
    /// </summary>
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

    #endregion

    #region Helpers

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

    #endregion
}
