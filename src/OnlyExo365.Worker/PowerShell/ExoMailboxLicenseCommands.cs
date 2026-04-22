using System.Management.Automation;
using OnlyExo365.Contracts;
using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

internal sealed class ExoMailboxLicenseCommands : ExoCommandModuleBase
{
    private const string LicenseWriteScope = "LicenseAssignment.ReadWrite.All";
    private static readonly string[] RequiredGraphModules =
    [
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Users",
        "Microsoft.Graph.Users.Actions",
        "Microsoft.Graph.Identity.DirectoryManagement"
    ];
    private readonly PowerShellEngine _engine;

    public ExoMailboxLicenseCommands(PowerShellEngine engine)
        : base(engine)
    {
        _engine = engine;
    }

    public async Task<List<TenantLicenseDto>> GetTenantLicensesAsync(CancellationToken cancellationToken = default)
    {
        await EnsureGraphConnectedAsync(requireLicenseWriteScope: false, cancellationToken);

        var script = BuildImportRequiredGraphModulesScript() + @"
try {
    $skus = Get-MgSubscribedSku -ErrorAction Stop
    foreach ($sku in $skus) {
        [PSCustomObject]@{
            SkuId = $sku.SkuId
            SkuPartNumber = $sku.SkuPartNumber
            ServicePlans = @($sku.ServicePlans | ForEach-Object {
                [PSCustomObject]@{
                    ServicePlanName = $_.ServicePlanName
                    ServicePlanId = $_.ServicePlanId
                    ProvisioningStatus = $_.ProvisioningStatus
                }
            })
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
            var skuId = GetString(obj, "SkuId");
            var skuPartNumber = GetString(obj, "SkuPartNumber");
            var servicePlans = MergeServicePlans(
                GetServicePlans(obj, "ServicePlans"),
                LicenseSkuNameResolver.GetServicePlans(skuPartNumber, skuId));
            licenses.Add(new TenantLicenseDto
            {
                SkuId = skuId,
                SkuPartNumber = skuPartNumber,
                DisplayName = LicenseSkuNameResolver.Resolve(skuPartNumber, skuId),
                ServicePlans = servicePlans,
                HasExchangeOnlineServicePlan = HasExchangeOnlineServicePlan(servicePlans),
                Total = GetInt(obj, "Total"),
                Assigned = GetInt(obj, "Assigned"),
                Available = GetInt(obj, "Available")
            });
        }

        return licenses;
    }

    public async Task<List<AdminRoleMemberDto>> GetAdminRoleMembersAsync(CancellationToken cancellationToken = default)
        => (await GetAdminRoleMembersDetailedAsync(cancellationToken).ConfigureAwait(false)).Members;

    internal async Task<GetAdminRoleMembersResult> GetAdminRoleMembersDetailedAsync(CancellationToken cancellationToken = default)
    {
        await EnsureGraphConnectedAsync(requireLicenseWriteScope: false, cancellationToken);

        var script = BuildGetAdminRoleMembersScript();
        var result = await _engine.ExecuteAsync(script, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        var response = new GetAdminRoleMembersResult
        {
            WarningDetails = ParseStructuredWarnings(result.Warning)
        };

        response.Warnings = ExtractWarningMessages(response.WarningDetails);

        foreach (var obj in result.Output)
        {
            response.Members.Add(new AdminRoleMemberDto
            {
                DisplayName = GetString(obj, "DisplayName"),
                UserPrincipalName = GetString(obj, "UserPrincipalName"),
                RoleName = GetString(obj, "RoleName")
            });
        }

        return response;
    }

    internal static string BuildGetAdminRoleMembersScript()
    {
        return BuildImportRequiredGraphModulesScript() + @"
function Write-AdminRoleWarning([string]$code, [string]$message, [string]$roleName) {
    $sampleItems = @()
    if (-not [string]::IsNullOrWhiteSpace($roleName)) {
        $sampleItems += $roleName
    }

    $warningPayload = @{
        Code = $code
        Scope = 'Dashboard.AdminUsers'
        Message = $message
        ItemIdentity = if ([string]::IsNullOrWhiteSpace($roleName)) { $null } else { $roleName }
        AffectedItemCount = if ($sampleItems.Count -gt 0) { $sampleItems.Count } else { $null }
        SampleItems = @($sampleItems)
        IsPartialData = $true
    } | ConvertTo-Json -Compress -Depth 4

    Write-Warning ""__EA_WARN__$warningPayload""
}

$adminRoles = @('Global Administrator', 'Exchange Administrator', 'User Administrator', 'Security Administrator', 'Helpdesk Administrator', 'SharePoint Administrator', 'Teams Administrator', 'Billing Administrator')

try {
    $directoryRoles = @(Get-MgDirectoryRole -All -ErrorAction Stop)

    foreach ($roleName in $adminRoles) {
        $role = @($directoryRoles | Where-Object { $_.DisplayName -eq $roleName } | Select-Object -First 1)
        if ($role.Count -eq 0) {
            continue
        }

        try {
            $members = @(Get-MgDirectoryRoleMember -DirectoryRoleId $role[0].Id -ErrorAction Stop)
        }
        catch {
            Write-AdminRoleWarning 'AdminRoleMembersLoadFailed' ""Get-MgDirectoryRoleMember failed for role '$roleName': $($_.Exception.Message)"" $roleName
            continue
        }

        foreach ($member in $members) {
            if ($member.AdditionalProperties.'@odata.type' -ne '#microsoft.graph.user') {
                continue
            }

            try {
                $user = Get-MgUser -UserId $member.Id -Property DisplayName,UserPrincipalName,AccountEnabled -ErrorAction Stop
            }
            catch {
                    Write-AdminRoleWarning 'AdminRoleUserLoadFailed' ""Get-MgUser failed for a member of role '$roleName': $($_.Exception.Message)"" $roleName
                continue
            }

            if ($user) {
                [PSCustomObject]@{
                    DisplayName = $user.DisplayName
                    UserPrincipalName = $user.UserPrincipalName
                    RoleName = $roleName
                    AccountEnabled = $user.AccountEnabled
                }
            }
        }
    }
} catch {
            Write-AdminRoleWarning 'AdminRoleDirectoryQueryFailed' ""Get-MgDirectoryRole failed: $($_.Exception.Message)"" ''
}";
    }

    public async Task<GetUserLicensesResponse> GetUserLicensesAsync(string userPrincipalName, CancellationToken cancellationToken = default)
    {
        await EnsureGraphConnectedAsync(requireLicenseWriteScope: false, cancellationToken);

        var script = BuildImportRequiredGraphModulesScript() + $@"
try {{
    $licenses = Get-MgUserLicenseDetail -UserId '{EscapePs(userPrincipalName)}' -ErrorAction Stop
    foreach ($lic in $licenses) {{
        [PSCustomObject]@{{
            SkuId = $lic.SkuId
            SkuPartNumber = $lic.SkuPartNumber
            ServicePlans = @($lic.ServicePlans | ForEach-Object {{
                [PSCustomObject]@{{
                    ServicePlanName = $_.ServicePlanName
                    ServicePlanId = $_.ServicePlanId
                    ProvisioningStatus = $_.ProvisioningStatus
                }}
            }})
        }}
    }}
}} catch {{
    Write-Warning ""Get-MgUserLicenseDetail not available: $($_.Exception.Message)""
}}";
        var results = await RunScriptAsync(script, cancellationToken);
        var licenses = new List<UserLicenseDto>();
        foreach (var obj in results)
        {
            var skuId = GetString(obj, "SkuId");
            var skuPartNumber = GetString(obj, "SkuPartNumber");
            var servicePlans = MergeServicePlans(
                GetServicePlans(obj, "ServicePlans"),
                LicenseSkuNameResolver.GetServicePlans(skuPartNumber, skuId));
            licenses.Add(new UserLicenseDto
            {
                SkuId = skuId,
                SkuPartNumber = skuPartNumber,
                DisplayName = LicenseSkuNameResolver.Resolve(skuPartNumber, skuId),
                ServicePlans = servicePlans,
                HasExchangeOnlineServicePlan = HasExchangeOnlineServicePlan(servicePlans)
            });
        }

        return new GetUserLicensesResponse { Licenses = licenses };
    }

    public async Task<GetMailboxProvisioningCandidatesResponse> GetMailboxProvisioningCandidatesAsync(
        GetMailboxProvisioningCandidatesRequest request,
        CancellationToken cancellationToken = default)
    {
        await EnsureGraphConnectedAsync(requireLicenseWriteScope: false, cancellationToken);

        var normalizedSearch = string.IsNullOrWhiteSpace(request.SearchQuery)
            ? null
            : request.SearchQuery.Trim();
        var pageSize = Math.Max(1, request.PageSize);
        var skip = Math.Max(0, request.Skip);

        var script = BuildImportRequiredGraphModulesScript() + $@"
$searchQuery = '{EscapePs(normalizedSearch)}'
$onlyWithoutLicense = {ToPsBoolLiteral(request.OnlyWithoutLicense)}
$onlyWithoutMail = {ToPsBoolLiteral(request.OnlyWithoutMail)}
$skip = {skip}
$pageSize = {pageSize}

$items = Get-MgUser -All -Filter ""userType eq 'Member'"" -Property DisplayName,UserPrincipalName,Mail,AccountEnabled,AssignedLicenses -ErrorAction Stop |
    ForEach-Object {{
        $licenseAssignments = @($_.AssignedLicenses)
        $hasAssignedLicense = $licenseAssignments.Count -gt 0
        $hasMailAddress = -not [string]::IsNullOrWhiteSpace($_.Mail)

        [PSCustomObject]@{{
            DisplayName = $_.DisplayName
            UserPrincipalName = $_.UserPrincipalName
            Mail = $_.Mail
            AccountEnabled = [bool]$_.AccountEnabled
            HasAssignedLicense = $hasAssignedLicense
            HasMailAddress = $hasMailAddress
        }}
    }}

if (-not [string]::IsNullOrWhiteSpace($searchQuery)) {{
    $items = $items | Where-Object {{
        ($_.DisplayName -like ""*$searchQuery*"") -or
        ($_.UserPrincipalName -like ""*$searchQuery*"") -or
        ($_.Mail -like ""*$searchQuery*"")
    }}
}}

if ($onlyWithoutLicense) {{
    $items = $items | Where-Object {{ -not $_.HasAssignedLicense }}
}}

if ($onlyWithoutMail) {{
    $items = $items | Where-Object {{ -not $_.HasMailAddress }}
}}

$items = @($items | Sort-Object DisplayName, UserPrincipalName)
$totalCount = $items.Count
$page = @($items | Select-Object -Skip $skip -First $pageSize)

[PSCustomObject]@{{
    TotalCount = $totalCount
    Skip = $skip
    PageSize = $pageSize
    HasMore = (($skip + $page.Count) -lt $totalCount)
    Candidates = $page
}}";

        var results = await RunScriptAsync(script, cancellationToken);
        var response = new GetMailboxProvisioningCandidatesResponse
        {
            SearchQuery = normalizedSearch,
            Skip = skip,
            PageSize = pageSize
        };

        var resultObject = results.FirstOrDefault();
        if (resultObject == null)
        {
            return response;
        }

        response.TotalCount = GetInt(resultObject, "TotalCount");
        response.Skip = GetInt(resultObject, "Skip");
        response.PageSize = GetInt(resultObject, "PageSize");
        response.HasMore = GetBool(resultObject, "HasMore");

        var candidatesValue = resultObject.Properties["Candidates"]?.Value;
        if (candidatesValue is IEnumerable<object> objects)
        {
            foreach (var entry in objects)
            {
                var candidate = entry as PSObject ?? PSObject.AsPSObject(entry);
                response.Candidates.Add(new MailboxProvisioningCandidateDto
                {
                    DisplayName = GetString(candidate, "DisplayName"),
                    UserPrincipalName = GetString(candidate, "UserPrincipalName"),
                    Mail = candidate.Properties["Mail"]?.Value?.ToString(),
                    AccountEnabled = GetBool(candidate, "AccountEnabled"),
                    HasAssignedLicense = GetBool(candidate, "HasAssignedLicense"),
                    HasMailAddress = GetBool(candidate, "HasMailAddress")
                });
            }
        }

        return response;
    }

    public async Task SetUserLicenseAsync(SetUserLicenseRequest request, CancellationToken cancellationToken = default)
    {
        await EnsureGraphConnectedAsync(requireLicenseWriteScope: true, cancellationToken);

        var script = BuildSetUserLicenseScript(request);
        await RunScriptAsync(script, cancellationToken);
    }

    public async Task<GetUsageLocationSuggestionResponse> GetUsageLocationSuggestionAsync(
        GetUsageLocationSuggestionRequest request,
        CancellationToken cancellationToken = default)
    {
        await EnsureGraphConnectedAsync(requireLicenseWriteScope: false, cancellationToken);

        var script = BuildGetUsageLocationSuggestionScript(request, _engine.ExchangeConfiguration.DefaultUsageLocation);
        var results = await RunScriptAsync(script, cancellationToken);
        var result = results.FirstOrDefault();
        if (result == null)
        {
            return new GetUsageLocationSuggestionResponse
            {
                UserPrincipalName = request.UserPrincipalName
            };
        }

        return new GetUsageLocationSuggestionResponse
        {
            UserPrincipalName = GetString(result, "UserPrincipalName"),
            CurrentUsageLocation = GetNullableString(result, "CurrentUsageLocation"),
            SuggestedUsageLocation = GetNullableString(result, "SuggestedUsageLocation"),
            SuggestionSource = GetNullableString(result, "SuggestionSource"),
            SuggestionDetails = GetNullableString(result, "SuggestionDetails")
        };
    }

    public async Task SetUserUsageLocationAsync(SetUserUsageLocationRequest request, CancellationToken cancellationToken = default)
    {
        await EnsureGraphConnectedAsync(requireLicenseWriteScope: true, cancellationToken);

        var script = BuildSetUserUsageLocationScript(request);
        await RunScriptAsync(script, cancellationToken);
    }

    internal static string BuildSetUserLicenseScript(SetUserLicenseRequest request)
    {
        var addSkus = string.Join(",", request.AddLicenseSkuIds.Select(s => $"@{{SkuId='{EscapePs(s)}'}}"));
        var removeSkus = string.Join(",", request.RemoveLicenseSkuIds.Select(s => $"'{EscapePs(s)}'"));
        var shouldValidateUsageLocation = request.AddLicenseSkuIds.Count > 0;

        return BuildImportRequiredGraphModulesScript() + $@"
$userId = '{EscapePs(request.UserPrincipalName)}'
$addLicenses = @({addSkus})
$removeLicenses = @({removeSkus})
$shouldValidateUsageLocation = {ToPsBoolLiteral(shouldValidateUsageLocation)}

if ($shouldValidateUsageLocation) {{
    $user = Get-MgUser -UserId $userId -Property Id,DisplayName,UserPrincipalName,UsageLocation -ErrorAction Stop
    $usageLocation = [string]$user.UsageLocation

    if ([string]::IsNullOrWhiteSpace($usageLocation)) {{
        throw ""License assignment cannot be completed because the user usage location is not set. Set UsageLocation to a valid two-letter country/region code and retry.""
    }}

    if ($usageLocation -notmatch '^[A-Za-z]{{2}}$') {{
        throw ""License assignment cannot be completed because the user usage location '$usageLocation' is invalid. Set UsageLocation to a valid two-letter country/region code and retry.""
    }}
}}

try {{
    Set-MgUserLicense -UserId $userId -AddLicenses $addLicenses -RemoveLicenses $removeLicenses -ErrorAction Stop
}}
catch {{
    $message = $_.Exception.Message
    if ($message -match 'invalid usage location') {{
        $currentUsageLocation = if ($shouldValidateUsageLocation -and $user) {{ [string]$user.UsageLocation }} else {{ '' }}
        $usageLocationSuffix = if ([string]::IsNullOrWhiteSpace($currentUsageLocation)) {{ '' }} else {{ "" Current value: '$currentUsageLocation'."" }}
        throw ""License assignment cannot be completed because the user usage location is missing or invalid.$usageLocationSuffix Set UsageLocation to a valid two-letter country/region code and retry.""
    }}

    throw
}}

Write-Output 'License updated successfully'";
    }

    internal static string BuildGetUsageLocationSuggestionScript(GetUsageLocationSuggestionRequest request, string? defaultUsageLocation)
    {
        var normalizedFallback = ExchangeOnlineConfiguration.NormalizeUsageLocation(defaultUsageLocation);

        return BuildImportRequiredGraphModulesScript() + $@"
$userId = '{EscapePs(request.UserPrincipalName)}'
$fallbackUsageLocation = '{EscapePs(normalizedFallback)}'

function Test-UsageLocationValue([string]$value) {{
    if ([string]::IsNullOrWhiteSpace($value)) {{
        return $false
    }}

    return $value -match '^[A-Za-z]{{2}}$'
}}

$user = Get-MgUser -UserId $userId -Property UserPrincipalName,UsageLocation -ErrorAction Stop
$organization = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1

$currentUsageLocation = [string]$user.UsageLocation
$tenantUsageLocation = if ($organization) {{ [string]$organization.CountryLetterCode }} else {{ '' }}
$suggestedUsageLocation = $null
$suggestionSource = $null
$suggestionDetails = $null

if (Test-UsageLocationValue $currentUsageLocation) {{
    $suggestedUsageLocation = $currentUsageLocation.ToUpperInvariant()
    $suggestionSource = 'User'
    $suggestionDetails = 'Existing UsageLocation already present on the user.'
}}
elseif (Test-UsageLocationValue $tenantUsageLocation) {{
    $suggestedUsageLocation = $tenantUsageLocation.ToUpperInvariant()
    $suggestionSource = 'Tenant'
    $suggestionDetails = 'Suggested from Microsoft Graph organization.countryLetterCode.'
}}
elseif (Test-UsageLocationValue $fallbackUsageLocation) {{
    $suggestedUsageLocation = $fallbackUsageLocation.ToUpperInvariant()
    $suggestionSource = 'Configuration'
    $suggestionDetails = 'Suggested from configured defaultUsageLocation fallback.'
}}
else {{
    $suggestionDetails = 'No valid UsageLocation could be derived from user, tenant or configuration fallback.'
}}

[PSCustomObject]@{{
    UserPrincipalName = $user.UserPrincipalName
    CurrentUsageLocation = if ([string]::IsNullOrWhiteSpace($currentUsageLocation)) {{ $null }} else {{ $currentUsageLocation.ToUpperInvariant() }}
    SuggestedUsageLocation = $suggestedUsageLocation
    SuggestionSource = $suggestionSource
    SuggestionDetails = $suggestionDetails
}}";
    }

    internal static string BuildSetUserUsageLocationScript(SetUserUsageLocationRequest request)
    {
        var normalizedUsageLocation = ExchangeOnlineConfiguration.NormalizeUsageLocation(request.UsageLocation);
        if (!ExchangeOnlineConfiguration.IsValidUsageLocation(normalizedUsageLocation))
        {
            throw new InvalidOperationException("UsageLocation must be a valid two-letter country/region code.");
        }

        return BuildImportRequiredGraphModulesScript() + $@"
$userId = '{EscapePs(request.UserPrincipalName)}'
$usageLocation = '{EscapePs(normalizedUsageLocation)}'
Update-MgUser -UserId $userId -UsageLocation $usageLocation -ErrorAction Stop
Write-Output 'Usage location updated successfully'";
    }

    public async Task<GetAvailableLicensesResponse> GetAvailableLicensesAsync(CancellationToken cancellationToken = default)
    {
        await EnsureGraphConnectedAsync(requireLicenseWriteScope: false, cancellationToken);

        var script = BuildImportRequiredGraphModulesScript() + @"
try {
    Get-MgSubscribedSku -ErrorAction Stop | Where-Object { ($_.PrepaidUnits.Enabled - $_.ConsumedUnits) -gt 0 } | ForEach-Object {
        [PSCustomObject]@{
            SkuId = $_.SkuId
            SkuPartNumber = $_.SkuPartNumber
            ServicePlans = @($_.ServicePlans | ForEach-Object {
                [PSCustomObject]@{
                    ServicePlanName = $_.ServicePlanName
                    ServicePlanId = $_.ServicePlanId
                    ProvisioningStatus = $_.ProvisioningStatus
                }
            })
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
            var skuId = GetString(obj, "SkuId");
            var skuPartNumber = GetString(obj, "SkuPartNumber");
            var servicePlans = MergeServicePlans(
                GetServicePlans(obj, "ServicePlans"),
                LicenseSkuNameResolver.GetServicePlans(skuPartNumber, skuId));
            licenses.Add(new TenantLicenseDto
            {
                SkuId = skuId,
                SkuPartNumber = skuPartNumber,
                DisplayName = LicenseSkuNameResolver.Resolve(skuPartNumber, skuId),
                ServicePlans = servicePlans,
                HasExchangeOnlineServicePlan = HasExchangeOnlineServicePlan(servicePlans),
                Total = GetInt(obj, "Total"),
                Assigned = GetInt(obj, "Assigned"),
                Available = GetInt(obj, "Available")
            });
        }

        return new GetAvailableLicensesResponse { Licenses = licenses };
    }

    private new async Task<List<PSObject>> RunScriptAsync(string script, CancellationToken cancellationToken)
    {
        var result = await _engine.ExecuteAsync(script, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException(result.ErrorMessage ?? "PowerShell command failed");
        }

        return result.Output;
    }

    private async Task<List<PSObject>> RunScriptAllowErrorsAsync(string script, CancellationToken cancellationToken)
    {
        var result = await _engine.ExecuteAsync(script, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        return result.Output;
    }

    private new static string GetString(PSObject obj, string propertyName)
    {
        var value = obj.Properties[propertyName]?.Value;
        return value?.ToString() ?? string.Empty;
    }

    private new static string? GetNullableString(PSObject obj, string propertyName)
    {
        var value = obj.Properties[propertyName]?.Value?.ToString();
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    private static int GetInt(PSObject obj, string propertyName)
    {
        var value = obj.Properties[propertyName]?.Value;
        return value switch
        {
            int intValue => intValue,
            long longValue => (int)longValue,
            _ when int.TryParse(value?.ToString(), out var parsed) => parsed,
            _ => 0
        };
    }

    private new static bool GetBool(PSObject obj, string propertyName)
    {
        var value = obj.Properties[propertyName]?.Value;
        return value switch
        {
            bool boolValue => boolValue,
            _ when bool.TryParse(value?.ToString(), out var parsed) => parsed,
            _ => false
        };
    }

    private static List<LicenseServicePlanDto> GetServicePlans(PSObject obj, string propertyName)
    {
        var value = obj.Properties[propertyName]?.Value;
        if (value is not System.Collections.IEnumerable enumerable || value is string)
        {
            return [];
        }

        var plans = new List<LicenseServicePlanDto>();
        foreach (var item in enumerable)
        {
            var plan = item as PSObject ?? PSObject.AsPSObject(item);
            var servicePlanName = GetString(plan, "ServicePlanName");
            var servicePlanId = GetString(plan, "ServicePlanId");
            if (string.IsNullOrWhiteSpace(servicePlanName) && string.IsNullOrWhiteSpace(servicePlanId))
            {
                continue;
            }

            plans.Add(new LicenseServicePlanDto
            {
                ServicePlanName = servicePlanName,
                ServicePlanId = servicePlanId,
                ProvisioningStatus = GetNullableString(plan, "ProvisioningStatus")
            });
        }

        return plans;
    }

    private static List<LicenseServicePlanDto> MergeServicePlans(
        IReadOnlyList<LicenseServicePlanDto> graphServicePlans,
        IReadOnlyList<LicenseSkuServicePlan> catalogServicePlans)
    {
        var merged = new List<LicenseServicePlanDto>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var plan in graphServicePlans)
        {
            var catalogMatch = catalogServicePlans.FirstOrDefault(item =>
                (!string.IsNullOrWhiteSpace(plan.ServicePlanId) &&
                 string.Equals(item.ServicePlanId, plan.ServicePlanId, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrWhiteSpace(plan.ServicePlanName) &&
                 string.Equals(item.ServicePlanName, plan.ServicePlanName, StringComparison.OrdinalIgnoreCase)));

            var key = BuildServicePlanKey(plan.ServicePlanName, plan.ServicePlanId);
            if (seen.Add(key))
            {
                merged.Add(new LicenseServicePlanDto
                {
                    ServicePlanName = plan.ServicePlanName,
                    ServicePlanId = plan.ServicePlanId,
                    FriendlyName = catalogMatch?.FriendlyName ?? plan.FriendlyName,
                    ProvisioningStatus = plan.ProvisioningStatus
                });
            }
        }

        foreach (var plan in catalogServicePlans)
        {
            var key = BuildServicePlanKey(plan.ServicePlanName, plan.ServicePlanId);
            if (seen.Add(key))
            {
                merged.Add(new LicenseServicePlanDto
                {
                    ServicePlanName = plan.ServicePlanName,
                    ServicePlanId = plan.ServicePlanId,
                    FriendlyName = plan.FriendlyName
                });
            }
        }

        return merged;
    }

    private static bool HasExchangeOnlineServicePlan(IEnumerable<LicenseServicePlanDto> servicePlans)
        => servicePlans.Any(plan =>
            string.Equals(plan.ServicePlanName, "EXCHANGE_S_STANDARD", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(plan.ServicePlanName, "EXCHANGE_S_ENTERPRISE", StringComparison.OrdinalIgnoreCase));

    private static string BuildServicePlanKey(string? servicePlanName, string? servicePlanId)
        => !string.IsNullOrWhiteSpace(servicePlanId)
            ? servicePlanId.Trim()
            : servicePlanName?.Trim() ?? string.Empty;

    private new static string EscapePs(string? value)
        => (value ?? string.Empty).Replace("'", "''");

    private new static string ToPsBoolLiteral(bool value)
        => value ? "$true" : "$false";

    internal static string BuildImportRequiredGraphModulesScript()
    {
        var moduleList = string.Join(", ", RequiredGraphModules.Select(module => $"'{EscapePs(module)}'"));
        return $@"
$requiredGraphModules = @({moduleList})
foreach ($requiredGraphModule in $requiredGraphModules) {{
    $availableModule = Get-Module -ListAvailable -Name $requiredGraphModule | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $availableModule) {{
        throw ""Required Microsoft Graph module '$requiredGraphModule' is not installed. Install the approved Graph bundle from Tools and retry.""
    }}

    Import-Module $requiredGraphModule -ErrorAction Stop | Out-Null
}}

";
    }

    private async Task EnsureGraphConnectedAsync(bool requireLicenseWriteScope, CancellationToken cancellationToken)
    {
        var requiredScopes = requireLicenseWriteScope
            ? GetLicenseWriteScopesOrThrow()
            : _engine.ExchangeConfiguration.NormalizeGraphScopes();

        var graphConnectionResult = await _engine.ConnectMicrosoftGraphAsync(
            ignoreAutoConnectConfiguration: true,
            delegatedScopes: requiredScopes,
            cancellationToken: cancellationToken);

        if (graphConnectionResult.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!_engine.IsGraphConnected)
        {
            throw new InvalidOperationException(
                graphConnectionResult.ErrorMessage ??
                (requireLicenseWriteScope
                    ? $"Microsoft Graph connection failed. Configure ONLYEXO365_GRAPH_LICENSE_WRITE_SCOPES with {LicenseWriteScope} to enable license assignment changes."
                    : "Microsoft Graph connection failed. Configure ONLYEXO365_GRAPH_SCOPES with the minimum scopes required for the requested Graph operation."));
        }
    }

    private IReadOnlyList<string> GetLicenseWriteScopesOrThrow()
    {
        var writeScopes = _engine.ExchangeConfiguration.NormalizeGraphLicenseWriteScopes();
        if (!writeScopes.Contains(LicenseWriteScope, StringComparer.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"Mailbox licensing write operations are disabled by configuration. Add {LicenseWriteScope} to graphLicenseWriteScopes or ONLYEXO365_GRAPH_LICENSE_WRITE_SCOPES to enable Set-MgUserLicense.");
        }

        return _engine.ExchangeConfiguration.GetGraphScopesForLicenseWrite();
    }
}

internal sealed class GetAdminRoleMembersResult
{
    public List<AdminRoleMemberDto> Members { get; } = new();

    public List<string> Warnings { get; set; } = new();

    public List<OperationWarningDto> WarningDetails { get; set; } = new();
}

