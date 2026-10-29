using System.Collections;
using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

internal sealed class ExoDashboardCommands : ExoCommandModuleBase
{
    private readonly CapabilityDetector _capabilityDetector;
    private readonly ExoMailboxLicenseCommands _mailboxLicenseCommands;

    public ExoDashboardCommands(
        PowerShellEngine engine,
        CapabilityDetector capabilityDetector,
        ExoMailboxLicenseCommands mailboxLicenseCommands)
        : base(engine)
    {
        _capabilityDetector = capabilityDetector;
        _mailboxLicenseCommands = mailboxLicenseCommands;
    }

    public async Task<DashboardStatsDto> GetDashboardStatsAsync(
        GetDashboardStatsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var stats = new DashboardStatsDto();
        var warnings = new List<string>();
        var warningDetails = new List<OperationWarningDto>();
        var quickCountLimit = stats.LargeTenantThreshold + 1;

        void AddWarning(string code, string message, string scope, bool isPartialData, int? affectedItemCount = null, IEnumerable<string>? sampleItems = null)
        {
            var warning = new OperationWarningDto
            {
                Code = code,
                Message = message,
                Scope = scope,
                IsPartialData = isPartialData,
                AffectedItemCount = affectedItemCount,
                SampleItems = sampleItems?
                    .Where(static item => !string.IsNullOrWhiteSpace(item))
                    .Take(5)
                    .ToList() ?? new List<string>()
            };

            warningDetails.Add(warning);
            warnings.Add(message);
        }

        onLog?.Invoke("Verbose", "Fetching mailbox counts...");

        var mailboxScript = request.QuickCount
            ? @"
$counts = @{
    UserMailboxes = 0
    SharedMailboxes = 0
    RoomMailboxes = 0
    EquipmentMailboxes = 0
    IsApproximate = $false
    FallbackTypes = @()
    UnavailableTypes = @()
}

$mailboxTypes = @('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox')
foreach ($mailboxType in $mailboxTypes) {
    try {
        $sample = @(Get-Mailbox -RecipientTypeDetails $mailboxType -ResultSize __LIMIT__ -ErrorAction Stop)
        $count = @($sample).Count
        if ($count -ge __LIMIT__) {
            $counts.IsApproximate = $true
            $count = __LIMIT__
        }

        switch ($mailboxType) {
            'UserMailbox' { $counts.UserMailboxes = $count }
            'SharedMailbox' { $counts.SharedMailboxes = $count }
            'RoomMailbox' { $counts.RoomMailboxes = $count }
            'EquipmentMailbox' { $counts.EquipmentMailboxes = $count }
        }
    }
    catch {
    try {
        $counts.IsApproximate = $true
        $counts.FallbackTypes += $mailboxType
        $fallback = @(Get-Mailbox -ResultSize 1000 -RecipientTypeDetails $mailboxType -ErrorAction Stop)
        $count = @($fallback).Count
        if ($count -ge 1000) {
            $counts.IsApproximate = $true
        }

        switch ($mailboxType) {
            'UserMailbox' { $counts.UserMailboxes = $count }
            'SharedMailbox' { $counts.SharedMailboxes = $count }
            'RoomMailbox' { $counts.RoomMailboxes = $count }
            'EquipmentMailbox' { $counts.EquipmentMailboxes = $count }
        }
    }
    catch {
        $counts.UnavailableTypes += $mailboxType
    }
    }
}

$counts
"
                .Replace("__LIMIT__", quickCountLimit.ToString())
            : @"
$counts = @{
    UserMailboxes = 0
    SharedMailboxes = 0
    RoomMailboxes = 0
    EquipmentMailboxes = 0
    IsApproximate = $false
    FallbackTypes = @()
    UnavailableTypes = @()
}

$mailboxTypes = @('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox')
foreach ($mailboxType in $mailboxTypes) {
    try {
        $count = @(Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails $mailboxType -ErrorAction Stop).Count

        switch ($mailboxType) {
            'UserMailbox' { $counts.UserMailboxes = $count }
            'SharedMailbox' { $counts.SharedMailboxes = $count }
            'RoomMailbox' { $counts.RoomMailboxes = $count }
            'EquipmentMailbox' { $counts.EquipmentMailboxes = $count }
        }
    }
    catch {
        try {
            if (-not (Get-Command -Name 'Get-EXOMailbox' -ErrorAction SilentlyContinue)) {
                throw 'Get-EXOMailbox not available'
            }

            $count = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails $mailboxType -ErrorAction Stop).Count

            switch ($mailboxType) {
                'UserMailbox' { $counts.UserMailboxes = $count }
                'SharedMailbox' { $counts.SharedMailboxes = $count }
                'RoomMailbox' { $counts.RoomMailboxes = $count }
                'EquipmentMailbox' { $counts.EquipmentMailboxes = $count }
            }
        }
        catch {
            $counts.UnavailableTypes += $mailboxType
        }
    }
}

$counts
";

        var mailboxResult = await Engine.ExecuteAsync(mailboxScript, onVerbose: onLog, cancellationToken: cancellationToken);

        if (mailboxResult.Success && mailboxResult.Output.Any())
        {
            var hash = mailboxResult.Output.First().BaseObject as Hashtable;
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

                var fallbackTypes = ConvertToStringList(hash["FallbackTypes"]);
                if (fallbackTypes.Count > 0)
                {
                    AddWarning(
                        code: "MailboxCountFallbackUsed",
                        message: $"Mailbox counts used limited fallback queries for: {string.Join(", ", fallbackTypes)}.",
                        scope: "Dashboard.MailboxCounts",
                        isPartialData: true,
                        affectedItemCount: fallbackTypes.Count,
                        sampleItems: fallbackTypes);
                }

                var unavailableTypes = ConvertToStringList(hash["UnavailableTypes"]);
                if (unavailableTypes.Count > 0)
                {
                    AddWarning(
                        code: "MailboxCountUnavailable",
                        message: $"Mailbox counts could not be retrieved for: {string.Join(", ", unavailableTypes)}.",
                        scope: "Dashboard.MailboxCounts",
                        isPartialData: true,
                        affectedItemCount: unavailableTypes.Count,
                        sampleItems: unavailableTypes);
                }
            }
        }
        else if (!mailboxResult.Success)
        {
            AddWarning(
                code: "MailboxCountsFailed",
                message: $"Mailbox counts could not be retrieved: {mailboxResult.ErrorMessage}",
                scope: "Dashboard.MailboxCounts",
                isPartialData: true);
        }

        cancellationToken.ThrowIfCancellationRequested();

        onLog?.Invoke("Verbose", "Fetching group counts...");

        var groupScript = request.QuickCount
            ? @"
$counts = @{
    DistributionGroups = 0
    DynamicDistributionGroups = 0
    IsApproximate = $false
    FallbackTypes = @()
    UnavailableTypes = @()
}

try {
    $distributionGroups = @(Get-DistributionGroup -ResultSize __LIMIT__ -ErrorAction Stop)
    $counts.DistributionGroups = @($distributionGroups).Count
    if ($counts.DistributionGroups -ge __LIMIT__) {
        $counts.IsApproximate = $true
    }
}
catch {
    $counts.IsApproximate = $true
    $counts.FallbackTypes += 'DistributionGroup'
    try { $counts.DistributionGroups = (Get-DistributionGroup -ResultSize 1000 -ErrorAction Stop).Count } catch { $counts.UnavailableTypes += 'DistributionGroup' }
}

try {
    $dynamicGroups = @(Get-DynamicDistributionGroup -ResultSize __LIMIT__ -ErrorAction Stop)
    $counts.DynamicDistributionGroups = @($dynamicGroups).Count
    if ($counts.DynamicDistributionGroups -ge __LIMIT__) {
        $counts.IsApproximate = $true
    }
}
catch {
    $counts.IsApproximate = $true
    $counts.FallbackTypes += 'DynamicDistributionGroup'
    try { $counts.DynamicDistributionGroups = (Get-DynamicDistributionGroup -ResultSize 1000 -ErrorAction Stop).Count } catch { $counts.UnavailableTypes += 'DynamicDistributionGroup' }
}

$counts
"
                .Replace("__LIMIT__", quickCountLimit.ToString())
            : @"
$counts = @{
    DistributionGroups = 0
    DynamicDistributionGroups = 0
    IsApproximate = $false
    FallbackTypes = @()
    UnavailableTypes = @()
}

try {
    $counts.DistributionGroups = (Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop).Count
}
catch {
    $counts.UnavailableTypes += 'DistributionGroup'
}

try {
    $counts.DynamicDistributionGroups = (Get-DynamicDistributionGroup -ResultSize Unlimited -ErrorAction Stop).Count
}
catch {
    $counts.UnavailableTypes += 'DynamicDistributionGroup'
}

$counts
";

        var groupResult = await Engine.ExecuteAsync(groupScript, onVerbose: onLog, cancellationToken: cancellationToken);

        if (groupResult.Success && groupResult.Output.Any())
        {
            var hash = groupResult.Output.First().BaseObject as Hashtable;
            if (hash != null)
            {
                stats.GroupCounts = new GroupCountsDto
                {
                    DistributionGroups = Convert.ToInt32(hash["DistributionGroups"] ?? 0),
                    DynamicDistributionGroups = Convert.ToInt32(hash["DynamicDistributionGroups"] ?? 0),
                    IsApproximate = hash["IsApproximate"] as bool? ?? false
                };

                var fallbackTypes = ConvertToStringList(hash["FallbackTypes"]);
                if (fallbackTypes.Count > 0)
                {
                    AddWarning(
                        code: "GroupCountFallbackUsed",
                        message: $"Group counts used limited fallback queries for: {string.Join(", ", fallbackTypes)}.",
                        scope: "Dashboard.GroupCounts",
                        isPartialData: true,
                        affectedItemCount: fallbackTypes.Count,
                        sampleItems: fallbackTypes);
                }

                var unavailableTypes = ConvertToStringList(hash["UnavailableTypes"]);
                if (unavailableTypes.Count > 0)
                {
                    AddWarning(
                        code: "GroupCountUnavailable",
                        message: $"Group counts could not be retrieved for: {string.Join(", ", unavailableTypes)}.",
                        scope: "Dashboard.GroupCounts",
                        isPartialData: true,
                        affectedItemCount: unavailableTypes.Count,
                        sampleItems: unavailableTypes);
                }
            }
        }
        else if (!groupResult.Success)
        {
            AddWarning(
                code: "GroupCountsFailed",
                message: $"Group counts could not be retrieved: {groupResult.ErrorMessage}",
                scope: "Dashboard.GroupCounts",
                isPartialData: true);
        }

        cancellationToken.ThrowIfCancellationRequested();

        var capabilities = _capabilityDetector.CachedCapabilities;
        if (request.IncludeUnifiedGroups && capabilities?.Features.CanGetUnifiedGroup == true)
        {
            onLog?.Invoke("Verbose", "Fetching unified group counts...");

            var unifiedScript = request.QuickCount
                ? @"
try {
    $count = @(Get-UnifiedGroup -ResultSize __LIMIT__).Count
    @{
        Count = $count
        Available = $true
        IsApproximate = $count -ge __LIMIT__
    }
}
catch {
    @{ Count = 0; Available = $false; Error = $_.Exception.Message; IsApproximate = $false }
}
"
                    .Replace("__LIMIT__", quickCountLimit.ToString())
                : @"
try {
    $count = (Get-UnifiedGroup -ResultSize Unlimited).Count
    @{ Count = $count; Available = $true; IsApproximate = $false }
}
catch {
    @{ Count = 0; Available = $false; Error = $_.Exception.Message; IsApproximate = $false }
}
";

            var unifiedResult = await Engine.ExecuteAsync(unifiedScript, onVerbose: onLog, cancellationToken: cancellationToken);

            if (unifiedResult.Success && unifiedResult.Output.Any())
            {
                var hash = unifiedResult.Output.First().BaseObject as Hashtable;
                if (hash != null)
                {
                    stats.GroupCounts.UnifiedGroupsAvailable = hash["Available"] as bool? ?? false;
                    if (stats.GroupCounts.UnifiedGroupsAvailable)
                    {
                        stats.GroupCounts.UnifiedGroups = Convert.ToInt32(hash["Count"] ?? 0);
                        stats.GroupCounts.IsApproximate |= hash["IsApproximate"] as bool? ?? false;
                    }
                    else
                    {
                        AddWarning(
                            code: "UnifiedGroupsUnavailable",
                            message: $"Unified Groups not available: {hash["Error"]}",
                            scope: "Dashboard.GroupCounts",
                            isPartialData: true);
                    }
                }
            }
            else if (!unifiedResult.Success)
            {
                AddWarning(
                    code: "UnifiedGroupsQueryFailed",
                    message: $"Unified Groups query failed: {unifiedResult.ErrorMessage}",
                    scope: "Dashboard.GroupCounts",
                    isPartialData: true);
            }
        }

        var totalMailboxes = stats.MailboxCounts.Total;
        if (totalMailboxes > stats.LargeTenantThreshold)
        {
            stats.IsLargeTenant = true;
            AddWarning(
                code: "LargeTenantDetected",
                message: $"Large tenant detected ({totalMailboxes} mailboxes). Some operations may be slower.",
                scope: "Dashboard",
                isPartialData: false);
        }

        cancellationToken.ThrowIfCancellationRequested();

        if (stats.MailboxCounts.IsApproximate)
        {
            AddWarning(
                code: "MailboxCountsApproximate",
                message: "Mailbox counts are approximate due to fallback limits.",
                scope: "Dashboard.MailboxCounts",
                isPartialData: true);
        }

        if (stats.GroupCounts.IsApproximate)
        {
            AddWarning(
                code: "GroupCountsApproximate",
                message: "Group counts are approximate due to quick-count limits.",
                scope: "Dashboard.GroupCounts",
                isPartialData: true);
        }

        onLog?.Invoke("Verbose", "Fetching tenant licenses...");
        try
        {
            stats.Licenses = await _mailboxLicenseCommands.GetTenantLicensesAsync(cancellationToken);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            onLog?.Invoke("Warning", $"Could not fetch tenant licenses: {ex.Message}");
            AddWarning(
                code: "TenantLicensesUnavailable",
                message: $"Could not fetch tenant licenses: {ex.Message}",
                scope: "Dashboard.Licenses",
                isPartialData: true);
        }

        cancellationToken.ThrowIfCancellationRequested();

        onLog?.Invoke("Verbose", "Fetching admin role members...");
        try
        {
            var adminRoleMembers = await _mailboxLicenseCommands.GetAdminRoleMembersDetailedAsync(cancellationToken);
            stats.AdminUsers = adminRoleMembers.Members;

            foreach (var warning in adminRoleMembers.WarningDetails)
            {
                AddWarning(
                    code: string.IsNullOrWhiteSpace(warning.Code) ? "AdminUsersPartialData" : warning.Code,
                    message: warning.Message,
                    scope: string.IsNullOrWhiteSpace(warning.Scope) ? "Dashboard.AdminUsers" : warning.Scope,
                    isPartialData: warning.IsPartialData,
                    affectedItemCount: warning.AffectedItemCount,
                    sampleItems: warning.SampleItems);
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            onLog?.Invoke("Warning", $"Could not fetch admin role members: {ex.Message}");
            AddWarning(
                code: "AdminUsersUnavailable",
                message: $"Could not fetch admin role members: {ex.Message}",
                scope: "Dashboard.AdminUsers",
                isPartialData: true);
        }

        stats.WarningDetails = warningDetails;
        stats.Warnings = warnings.Distinct(StringComparer.Ordinal).ToList();
        stats.HasPartialData = warningDetails.Any(static warning => warning.IsPartialData);
        stats.RetrievedAt = DateTime.UtcNow;

        return stats;
    }
}

