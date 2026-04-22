using System.Collections;
using System.Management.Automation;
using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

internal sealed class ExoMobileDeviceCommands : ExoCommandModuleBase
{
    private readonly CapabilityDetector _capabilityDetector;

    public ExoMobileDeviceCommands(PowerShellEngine engine, CapabilityDetector capabilityDetector)
        : base(engine)
    {
        _capabilityDetector = capabilityDetector;
    }

    public async Task<GetMobileDevicesResponse> GetMobileDevicesAsync(
        GetMobileDevicesRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var capabilities = await _capabilityDetector.DetectCapabilitiesAsync(cancellationToken: cancellationToken);
        MobileDeviceCapabilityGuard.EnsureListingAvailable(capabilities);

        var normalizedAccessState = NormalizeMobileDeviceAccessState(request.AccessState);
        var sortProperty = NormalizeMobileDeviceSortProperty(request.SortBy);

        var response = new GetMobileDevicesResponse
        {
            Skip = request.Skip,
            PageSize = request.PageSize,
            SearchQuery = request.SearchQuery
        };

        if (!string.IsNullOrWhiteSpace(request.AccessState) && normalizedAccessState == null)
        {
            onLog?.Invoke("Warning", $"Unsupported AccessState ignored: {request.AccessState}");
        }

        if (!string.IsNullOrWhiteSpace(request.SortBy) &&
            !string.Equals(sortProperty, request.SortBy.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            onLog?.Invoke("Warning", $"Unsupported SortBy ignored: {request.SortBy}");
        }

        var escapedSearch = EscapePs(request.SearchQuery);
        var escapedAccessState = EscapePs(normalizedAccessState);
        var sortDirection = request.SortDescending ? "-Descending" : string.Empty;

        var script = BuildGetMobileDevicesScript(
            request.Skip,
            request.PageSize,
            escapedSearch,
            escapedAccessState,
            sortProperty,
            sortDirection);

        onLog?.Invoke("Verbose", $"Fetching mobile devices (skip={request.Skip}, pageSize={request.PageSize}, accessState={normalizedAccessState ?? "All"})...");

        var result = await Engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (result.Success && result.Output.Any() && result.Output.First().BaseObject is Hashtable hash)
        {
            response.TotalCount = Convert.ToInt32(hash["TotalCount"] ?? 0);
            response.HasMore = hash["HasMore"] is bool hasMore
                ? hasMore
                : Convert.ToBoolean(hash["HasMore"] ?? false);
            response.IsTotalCountExact = hash["IsTotalCountExact"] is bool isTotalCountExact
                ? isTotalCountExact
                : Convert.ToBoolean(hash["IsTotalCountExact"] ?? true);

            if (hash["Devices"] is object[] devices)
            {
                foreach (var deviceObject in devices)
                {
                    if (deviceObject is not PSObject devicePs)
                    {
                        continue;
                    }

                    response.Devices.Add(new MobileDeviceListItemDto
                    {
                        Identity = GetString(devicePs, "Identity"),
                        Guid = GetNullableString(devicePs, "Guid"),
                        DeviceId = GetString(devicePs, "DeviceId"),
                        FriendlyName = GetNullableString(devicePs, "FriendlyName"),
                        DeviceType = GetNullableString(devicePs, "DeviceType"),
                        DeviceModel = GetNullableString(devicePs, "DeviceModel"),
                        DeviceUserAgent = GetNullableString(devicePs, "DeviceUserAgent"),
                        DeviceOS = GetNullableString(devicePs, "DeviceOS"),
                        ClientType = GetNullableString(devicePs, "ClientType"),
                        UserDisplayName = GetNullableString(devicePs, "UserDisplayName"),
                        UserPrincipalName = GetNullableString(devicePs, "UserPrincipalName"),
                        MailboxIdentity = GetString(devicePs, "MailboxIdentity"),
                        MailboxDisplayName = GetString(devicePs, "MailboxDisplayName"),
                        CurrentMailboxPolicy = GetNullableString(devicePs, "CurrentMailboxPolicy"),
                        DeviceAccessState = GetString(devicePs, "DeviceAccessState"),
                        DeviceAccessStateReason = GetNullableString(devicePs, "DeviceAccessStateReason"),
                        FirstSyncTime = GetNullableDateTime(devicePs, "FirstSyncTime"),
                        LastSuccessSync = GetNullableDateTime(devicePs, "LastSuccessSync"),
                        LastPolicyUpdateTime = GetNullableDateTime(devicePs, "LastPolicyUpdateTime"),
                        Status = GetNullableString(devicePs, "Status")
                    });
                }
            }
        }

        onLog?.Invoke("Information", $"Retrieved {response.Devices.Count} mobile devices (total: {response.TotalCount})");

        return response;
    }

    internal static string BuildGetMobileDevicesScript(
        int skip,
        int pageSize,
        string escapedSearch,
        string escapedAccessState,
        string sortProperty,
        string sortDirection)
    {
        return $@"
function Resolve-MailboxIdentity([string]$identity) {{
    if ([string]::IsNullOrWhiteSpace($identity)) {{
        return ''
    }}

    $marker = '\ExchangeActiveSyncDevices\'
    $index = $identity.IndexOf($marker, [System.StringComparison]::OrdinalIgnoreCase)
    if ($index -gt 0) {{
        return $identity.Substring(0, $index)
    }}

    return $identity
}}

function Resolve-UserPrincipalName([object]$source, [string]$mailboxIdentity) {{
    if ($null -ne $source) {{
        foreach ($propertyName in @('UserPrincipalName', 'MicrosoftOnlineServicesID', 'WindowsEmailAddress', 'PrimarySmtpAddress', 'User')) {{
            $property = $source.PSObject.Properties[$propertyName]
            if ($null -eq $property -or $null -eq $property.Value) {{
                continue
            }}

            $value = $property.Value.ToString().Trim()
            if ([string]::IsNullOrWhiteSpace($value)) {{
                continue
            }}

            if ($value.Contains('@') -or $propertyName -eq 'UserPrincipalName') {{
                return $value
            }}
        }}
    }}

    if (-not [string]::IsNullOrWhiteSpace($mailboxIdentity) -and $mailboxIdentity.Contains('@')) {{
        return $mailboxIdentity
    }}

    return $null
}}

function Resolve-PreferredMailboxIdentity([string]$mailboxIdentity, [string]$userPrincipalName) {{
    if (-not [string]::IsNullOrWhiteSpace($userPrincipalName) -and $userPrincipalName.Contains('@')) {{
        return $userPrincipalName.Trim()
    }}

    return $mailboxIdentity
}}

function Resolve-DisplayedMailboxLabel([string]$mailboxIdentity, [string]$userDisplayName, [string]$userPrincipalName) {{
    if (-not [string]::IsNullOrWhiteSpace($userDisplayName)) {{
        return $userDisplayName.Trim()
    }}

    if (-not [string]::IsNullOrWhiteSpace($userPrincipalName) -and $userPrincipalName.Contains('@')) {{
        return $userPrincipalName.Trim()
    }}

    return $mailboxIdentity
}}

$searchQuery = '{escapedSearch}'
$accessState = '{escapedAccessState}'

$items = Get-MobileDevice -ResultSize Unlimited -ErrorAction Stop | ForEach-Object {{
    $device = $_
    $mailboxIdentity = Resolve-MailboxIdentity $device.Identity.ToString()
    $userPrincipalName = Resolve-UserPrincipalName $device $mailboxIdentity
    $resolvedMailboxIdentity = Resolve-PreferredMailboxIdentity $mailboxIdentity $userPrincipalName
    $resolvedMailboxDisplayName = Resolve-DisplayedMailboxLabel $resolvedMailboxIdentity $device.UserDisplayName $userPrincipalName

    [PSCustomObject]@{{
        Identity = $device.Identity.ToString()
        Guid = if ($device.Guid) {{ $device.Guid.ToString() }} else {{ $null }}
        DeviceId = if ($device.DeviceId) {{ $device.DeviceId.ToString() }} else {{ '' }}
        FriendlyName = if ($device.FriendlyName) {{ $device.FriendlyName.ToString() }} else {{ $null }}
        DeviceType = if ($device.DeviceType) {{ $device.DeviceType.ToString() }} else {{ $null }}
        DeviceModel = if ($device.DeviceModel) {{ $device.DeviceModel.ToString() }} else {{ $null }}
        DeviceUserAgent = if ($device.DeviceUserAgent) {{ $device.DeviceUserAgent.ToString() }} else {{ $null }}
        DeviceOS = if ($device.DeviceOS) {{ $device.DeviceOS.ToString() }} else {{ $null }}
        ClientType = if ($device.ClientType) {{ $device.ClientType.ToString() }} else {{ $null }}
        UserDisplayName = if ($device.UserDisplayName) {{ $device.UserDisplayName.ToString() }} else {{ $null }}
        UserPrincipalName = $userPrincipalName
        MailboxIdentity = $resolvedMailboxIdentity
        MailboxDisplayName = $resolvedMailboxDisplayName
        CurrentMailboxPolicy = $null
        DeviceAccessState = if ($device.DeviceAccessState) {{ $device.DeviceAccessState.ToString() }} else {{ 'Unknown' }}
        DeviceAccessStateReason = if ($device.DeviceAccessStateReason) {{ $device.DeviceAccessStateReason.ToString() }} else {{ $null }}
        FirstSyncTime = $null
        LastSuccessSync = $null
        LastPolicyUpdateTime = $null
        Status = $null
    }}
}} | Where-Object {{
    $matchesSearch = $true
    if (-not [string]::IsNullOrWhiteSpace($searchQuery)) {{
        $matchesSearch =
            $_.UserDisplayName -like ""*$searchQuery*"" -or
            $_.UserPrincipalName -like ""*$searchQuery*"" -or
            $_.MailboxIdentity -like ""*$searchQuery*"" -or
            $_.DeviceId -like ""*$searchQuery*"" -or
            $_.FriendlyName -like ""*$searchQuery*"" -or
            $_.DeviceType -like ""*$searchQuery*"" -or
            $_.DeviceModel -like ""*$searchQuery*""
    }}

    if (-not $matchesSearch) {{
        return $false
    }}

    if (-not [string]::IsNullOrWhiteSpace($accessState) -and $_.DeviceAccessState -ne $accessState) {{
        return $false
    }}

    return $true
}} | Sort-Object {sortProperty} {sortDirection}

$allItems = @($items)
$totalCount = $allItems.Count
$pagedItems = @($allItems | Select-Object -Skip {skip} -First {pageSize})
$hasMore = ({skip} + $pagedItems.Count) -lt $totalCount

@{{
    TotalCount = $totalCount
    HasMore = $hasMore
    IsTotalCountExact = $true
    Devices = @($pagedItems)
}}";
    }

    internal static int CalculateMobileDevicePageWindowSize(int skip, int pageSize)
        => Math.Max(0, skip) + Math.Max(1, pageSize) + 1;

    public async Task<GetMobileDeviceDetailsResponse> GetMobileDeviceDetailsAsync(
        GetMobileDeviceDetailsRequest request,
        Action<string, string>? onLog = null,
        Action<int, int, string>? onProgress = null,
        CancellationToken cancellationToken = default)
    {
        var capabilities = await _capabilityDetector.DetectCapabilitiesAsync(cancellationToken: cancellationToken);
        MobileDeviceCapabilityGuard.EnsureListingAvailable(capabilities);

        var canGetStatistics = capabilities.Features.CanGetMobileDeviceStatistics;
        var canGetCasMailbox = capabilities.Features.CanGetCasMailbox;
        var totalSteps = 1 + (canGetStatistics ? 1 : 0) + (canGetCasMailbox && !string.IsNullOrWhiteSpace(request.MailboxIdentity) ? 1 : 0);
        var currentStep = 0;

        onLog?.Invoke("Verbose", $"Fetching mobile device details for {request.Identity}...");

        var deviceResults = await RunScriptAsync(BuildGetMobileDeviceDetailsScript(request.Identity), cancellationToken);
        var device = MapMobileDevice(deviceResults.FirstOrDefault())
            ?? throw new InvalidOperationException($"Mobile device not found: {request.Identity}");

        currentStep++;
        onProgress?.Invoke(currentStep, totalSteps, "Base mobile device data loaded");

        if (canGetStatistics)
        {
            var statisticsResults = await RunScriptAllowErrorsAsync(
                BuildGetMobileDeviceStatisticsScript(device.Identity),
                cancellationToken: cancellationToken);
            ApplyMobileDeviceStatistics(device, statisticsResults.FirstOrDefault());

            currentStep++;
            onProgress?.Invoke(currentStep, totalSteps, "Mobile device statistics loaded");
        }

        if (canGetCasMailbox && !string.IsNullOrWhiteSpace(device.MailboxIdentity))
        {
            var casResults = await RunScriptAllowErrorsAsync(
                BuildGetMobileDeviceCasMailboxScript(device.MailboxIdentity),
                cancellationToken: cancellationToken);
            ApplyMobileDeviceCasMailbox(device, casResults.FirstOrDefault());

            currentStep++;
            onProgress?.Invoke(currentStep, totalSteps, "Mailbox policy data loaded");
        }

        return new GetMobileDeviceDetailsResponse
        {
            Device = device
        };
    }

    internal static string BuildGetMobileDeviceDetailsScript(string identity)
    {
        var escapedIdentity = EscapePs(identity);

        return $@"
function Resolve-MailboxIdentity([string]$deviceIdentity) {{
    if ([string]::IsNullOrWhiteSpace($deviceIdentity)) {{
        return ''
    }}

    $marker = '\ExchangeActiveSyncDevices\'
    $index = $deviceIdentity.IndexOf($marker, [System.StringComparison]::OrdinalIgnoreCase)
    if ($index -gt 0) {{
        return $deviceIdentity.Substring(0, $index)
    }}

    return $deviceIdentity
}}

function Resolve-UserPrincipalName([object]$source, [string]$mailboxIdentity) {{
    if ($null -ne $source) {{
        foreach ($propertyName in @('UserPrincipalName', 'MicrosoftOnlineServicesID', 'WindowsEmailAddress', 'PrimarySmtpAddress', 'User')) {{
            $property = $source.PSObject.Properties[$propertyName]
            if ($null -eq $property -or $null -eq $property.Value) {{
                continue
            }}

            $value = $property.Value.ToString().Trim()
            if ([string]::IsNullOrWhiteSpace($value)) {{
                continue
            }}

            if ($value.Contains('@') -or $propertyName -eq 'UserPrincipalName') {{
                return $value
            }}
        }}
    }}

    if (-not [string]::IsNullOrWhiteSpace($mailboxIdentity) -and $mailboxIdentity.Contains('@')) {{
        return $mailboxIdentity
    }}

    return $null
}}

function Resolve-PreferredMailboxIdentity([string]$mailboxIdentity, [string]$userPrincipalName) {{
    if (-not [string]::IsNullOrWhiteSpace($userPrincipalName) -and $userPrincipalName.Contains('@')) {{
        return $userPrincipalName.Trim()
    }}

    return $mailboxIdentity
}}

function Resolve-DisplayedMailboxLabel([string]$mailboxIdentity, [string]$userDisplayName, [string]$userPrincipalName) {{
    if (-not [string]::IsNullOrWhiteSpace($userDisplayName)) {{
        return $userDisplayName.Trim()
    }}

    if (-not [string]::IsNullOrWhiteSpace($userPrincipalName) -and $userPrincipalName.Contains('@')) {{
        return $userPrincipalName.Trim()
    }}

    return $mailboxIdentity
}}

$device = Get-MobileDevice -Identity '{escapedIdentity}' -ErrorAction Stop | Select-Object -First 1
$mailboxIdentity = Resolve-MailboxIdentity $device.Identity.ToString()
$userPrincipalName = Resolve-UserPrincipalName $device $mailboxIdentity
$resolvedMailboxIdentity = Resolve-PreferredMailboxIdentity $mailboxIdentity $userPrincipalName
$resolvedMailboxDisplayName = Resolve-DisplayedMailboxLabel $resolvedMailboxIdentity $device.UserDisplayName $userPrincipalName

[PSCustomObject]@{{
    Identity = $device.Identity.ToString()
    Guid = if ($device.Guid) {{ $device.Guid.ToString() }} else {{ $null }}
    DeviceId = if ($device.DeviceId) {{ $device.DeviceId.ToString() }} else {{ '' }}
    FriendlyName = if ($device.FriendlyName) {{ $device.FriendlyName.ToString() }} else {{ $null }}
    DeviceType = if ($device.DeviceType) {{ $device.DeviceType.ToString() }} else {{ $null }}
    DeviceModel = if ($device.DeviceModel) {{ $device.DeviceModel.ToString() }} else {{ $null }}
    DeviceUserAgent = if ($device.DeviceUserAgent) {{ $device.DeviceUserAgent.ToString() }} else {{ $null }}
    DeviceOS = if ($device.DeviceOS) {{ $device.DeviceOS.ToString() }} else {{ $null }}
    ClientType = if ($device.ClientType) {{ $device.ClientType.ToString() }} else {{ $null }}
    UserDisplayName = if ($device.UserDisplayName) {{ $device.UserDisplayName.ToString() }} else {{ $null }}
    UserPrincipalName = $userPrincipalName
    MailboxIdentity = $resolvedMailboxIdentity
    MailboxDisplayName = $resolvedMailboxDisplayName
    CurrentMailboxPolicy = $null
    DeviceAccessState = if ($device.DeviceAccessState) {{ $device.DeviceAccessState.ToString() }} else {{ 'Unknown' }}
    DeviceAccessStateReason = if ($device.DeviceAccessStateReason) {{ $device.DeviceAccessStateReason.ToString() }} else {{ $null }}
    FirstSyncTime = $null
    LastSuccessSync = $null
    LastPolicyUpdateTime = $null
    Status = $null
}}";
    }

    internal static string BuildGetMobileDeviceStatisticsScript(string identity)
    {
        var escapedIdentity = EscapePs(identity);

        return $@"
$stats = Get-MobileDeviceStatistics -Identity '{escapedIdentity}' -ErrorAction Stop

[PSCustomObject]@{{
    FriendlyName = if ($stats.DeviceFriendlyName) {{ $stats.DeviceFriendlyName.ToString() }} else {{ $null }}
    DeviceUserAgent = if ($stats.DeviceUserAgent) {{ $stats.DeviceUserAgent.ToString() }} else {{ $null }}
    DeviceOS = if ($stats.DeviceOS) {{ $stats.DeviceOS.ToString() }} else {{ $null }}
    UserDisplayName = if ($stats.UserDisplayName) {{ $stats.UserDisplayName.ToString() }} else {{ $null }}
    UserPrincipalName = if ($stats.UserPrincipalName) {{ $stats.UserPrincipalName.ToString() }} else {{ $null }}
    FirstSyncTime = if ($stats.FirstSyncTime) {{ $stats.FirstSyncTime }} else {{ $null }}
    LastSuccessSync = if ($stats.LastSuccessSync) {{ $stats.LastSuccessSync }} else {{ $null }}
    LastPolicyUpdateTime = if ($stats.LastPolicyUpdateTime) {{ $stats.LastPolicyUpdateTime }} else {{ $null }}
    Status = if ($stats.Status) {{ $stats.Status.ToString() }} else {{ $null }}
}}";
    }

    internal static string BuildGetMobileDeviceCasMailboxScript(string mailboxIdentity)
    {
        var escapedMailboxIdentity = EscapePs(mailboxIdentity);

        return $@"
function Resolve-UserPrincipalName([object]$source, [string]$mailboxIdentity) {{
    if ($null -ne $source) {{
        foreach ($propertyName in @('UserPrincipalName', 'MicrosoftOnlineServicesID', 'WindowsEmailAddress', 'PrimarySmtpAddress', 'User')) {{
            $property = $source.PSObject.Properties[$propertyName]
            if ($null -eq $property -or $null -eq $property.Value) {{
                continue
            }}

            $value = $property.Value.ToString().Trim()
            if ([string]::IsNullOrWhiteSpace($value)) {{
                continue
            }}

            if ($value.Contains('@') -or $propertyName -eq 'UserPrincipalName') {{
                return $value
            }}
        }}
    }}

    if (-not [string]::IsNullOrWhiteSpace($mailboxIdentity) -and $mailboxIdentity.Contains('@')) {{
        return $mailboxIdentity
    }}

    return $null
}}

function Resolve-PreferredMailboxIdentity([string]$mailboxIdentity, [string]$userPrincipalName) {{
    if (-not [string]::IsNullOrWhiteSpace($userPrincipalName) -and $userPrincipalName.Contains('@')) {{
        return $userPrincipalName.Trim()
    }}

    return $mailboxIdentity
}}

$cas = Get-CASMailbox -Identity '{escapedMailboxIdentity}' -ErrorAction Stop
$userPrincipalName = Resolve-UserPrincipalName $cas '{escapedMailboxIdentity}'
$resolvedMailboxIdentity = Resolve-PreferredMailboxIdentity '{escapedMailboxIdentity}' $userPrincipalName

[PSCustomObject]@{{
    CurrentMailboxPolicy = if ($cas.ActiveSyncMailboxPolicy) {{ $cas.ActiveSyncMailboxPolicy.ToString() }} else {{ $null }}
    UserPrincipalName = $userPrincipalName
    MailboxIdentity = $resolvedMailboxIdentity
    MailboxDisplayName = if ($cas.DisplayName) {{ $cas.DisplayName.ToString() }} else {{ $null }}
}}";
    }

    public async Task<GetMobileDeviceMailboxPoliciesResponse> GetMobileDeviceMailboxPoliciesAsync(
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var capabilities = await _capabilityDetector.DetectCapabilitiesAsync(cancellationToken: cancellationToken);
        MobileDeviceCapabilityGuard.EnsurePoliciesAvailable(capabilities);

        var script = BuildGetMobileDeviceMailboxPoliciesScript();

        onLog?.Invoke("Verbose", "Fetching mobile device mailbox policies...");
        var results = await RunScriptAsync(script, cancellationToken);
        var response = new GetMobileDeviceMailboxPoliciesResponse();

        foreach (var obj in results)
        {
            response.Policies.Add(new MobileDeviceMailboxPolicyDto
            {
                Identity = GetString(obj, "Identity"),
                Name = GetString(obj, "Name"),
                IsDefault = GetBool(obj, "IsDefault"),
                PasswordEnabled = GetNullableBool(obj, "PasswordEnabled"),
                AlphanumericPasswordRequired = GetNullableBool(obj, "AlphanumericPasswordRequired"),
                AllowNonProvisionableDevices = GetNullableBool(obj, "AllowNonProvisionableDevices"),
                DeviceEncryptionEnabled = GetNullableBool(obj, "DeviceEncryptionEnabled"),
                AttachmentsEnabled = GetNullableBool(obj, "AttachmentsEnabled"),
                MaxAttachmentSize = GetNullableString(obj, "MaxAttachmentSize")
            });
        }

        return response;
    }

    internal static string BuildGetMobileDeviceMailboxPoliciesScript() => @"
Get-MobileDeviceMailboxPolicy -ErrorAction Stop | Sort-Object Name | ForEach-Object {
    [PSCustomObject]@{
        Identity = $_.Identity.ToString()
        Name = $_.Name
        IsDefault = [bool]$_.IsDefault
        PasswordEnabled = if ($null -ne $_.PasswordEnabled) { [bool]$_.PasswordEnabled } else { $null }
        AlphanumericPasswordRequired = if ($null -ne $_.AlphanumericPasswordRequired) { [bool]$_.AlphanumericPasswordRequired } else { $null }
        AllowNonProvisionableDevices = if ($null -ne $_.AllowNonProvisionableDevices) { [bool]$_.AllowNonProvisionableDevices } else { $null }
        DeviceEncryptionEnabled = if ($null -ne $_.DeviceEncryptionEnabled) { [bool]$_.DeviceEncryptionEnabled } else { $null }
        AttachmentsEnabled = if ($null -ne $_.AttachmentsEnabled) { [bool]$_.AttachmentsEnabled } else { $null }
        MaxAttachmentSize = if ($_.MaxAttachmentSize) { $_.MaxAttachmentSize.ToString() } else { $null }
    }
}";

    public async Task SetMobileDeviceAccessStateAsync(
        SetMobileDeviceAccessStateRequest request,
        CancellationToken cancellationToken = default)
    {
        var capabilities = await _capabilityDetector.DetectCapabilitiesAsync(cancellationToken: cancellationToken);
        MobileDeviceCapabilityGuard.EnsureAccessStateManagementAvailable(capabilities);

        var accessState = NormalizeMobileDeviceAccessState(request.AccessState)
            ?? throw new InvalidOperationException($"Unsupported mobile device access state: {request.AccessState}");

        var mailboxIdentity = EscapePs(request.MailboxIdentity);
        var deviceId = EscapePs(request.DeviceId);

        var script = $@"
$cas = Get-CASMailbox -Identity '{mailboxIdentity}' -ErrorAction Stop
$allowed = @($cas.ActiveSyncAllowedDeviceIDs | ForEach-Object {{ $_.ToString() }} | Where-Object {{ -not [string]::IsNullOrWhiteSpace($_) }})
$blocked = @($cas.ActiveSyncBlockedDeviceIDs | ForEach-Object {{ $_.ToString() }} | Where-Object {{ -not [string]::IsNullOrWhiteSpace($_) }})
$deviceId = '{deviceId}'

switch ('{accessState}') {{
    'Allowed' {{
        $allowed = @($allowed + $deviceId | Select-Object -Unique)
        $blocked = @($blocked | Where-Object {{ $_ -ne $deviceId }})
    }}
    'Blocked' {{
        $blocked = @($blocked + $deviceId | Select-Object -Unique)
        $allowed = @($allowed | Where-Object {{ $_ -ne $deviceId }})
    }}
    default {{
        $allowed = @($allowed | Where-Object {{ $_ -ne $deviceId }})
        $blocked = @($blocked | Where-Object {{ $_ -ne $deviceId }})
    }}
}}

Set-CASMailbox -Identity '{mailboxIdentity}' -ActiveSyncAllowedDeviceIDs $allowed -ActiveSyncBlockedDeviceIDs $blocked -ErrorAction Stop
Write-Output 'OK'";

        await RunScriptAsync(script, cancellationToken);
    }

    public async Task ClearMobileDeviceAsync(ClearMobileDeviceRequest request, CancellationToken cancellationToken = default)
    {
        var capabilities = await _capabilityDetector.DetectCapabilitiesAsync(cancellationToken: cancellationToken);
        MobileDeviceCapabilityGuard.EnsureRemoteWipeAvailable(capabilities);

        var identity = EscapePs(request.Identity);
        var script = $@"
Clear-MobileDevice -Identity '{identity}' -Confirm:$false -ErrorAction Stop
Write-Output 'OK'";

        await RunScriptAsync(script, cancellationToken);
    }

    public async Task SetMobileDeviceMailboxPolicyAsync(SetMobileDeviceMailboxPolicyRequest request, CancellationToken cancellationToken = default)
    {
        var capabilities = await _capabilityDetector.DetectCapabilitiesAsync(cancellationToken: cancellationToken);
        MobileDeviceCapabilityGuard.EnsureMailboxPolicyAssignmentAvailable(capabilities);

        var mailboxIdentity = EscapePs(request.MailboxIdentity);
        var policyIdentity = EscapePs(request.PolicyIdentity);
        var script = string.IsNullOrWhiteSpace(request.PolicyIdentity)
            ? $@"
Set-CASMailbox -Identity '{mailboxIdentity}' -ActiveSyncMailboxPolicy $null -ErrorAction Stop
Write-Output 'OK'"
            : $@"
Set-CASMailbox -Identity '{mailboxIdentity}' -ActiveSyncMailboxPolicy '{policyIdentity}' -ErrorAction Stop
Write-Output 'OK'";

        await RunScriptAsync(script, cancellationToken);
    }

    private static string? NormalizeMobileDeviceAccessState(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return value.Trim() switch
        {
            "Allowed" => "Allowed",
            "Blocked" => "Blocked",
            "Quarantined" => "Quarantined",
            "DeviceDiscovery" => "DeviceDiscovery",
            "Unknown" => "Unknown",
            _ => null
        };
    }

    private static string NormalizeMobileDeviceSortProperty(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "UserDisplayName";
        }

        return value.Trim() switch
        {
            "DeviceAccessState" => "DeviceAccessState",
            "LastSuccessSync" => "LastSuccessSync",
            "DeviceType" => "DeviceType",
            "DeviceModel" => "DeviceModel",
            "DeviceId" => "DeviceId",
            "CurrentMailboxPolicy" => "CurrentMailboxPolicy",
            _ => "UserDisplayName"
        };
    }

    private static MobileDeviceListItemDto? MapMobileDevice(PSObject? obj)
    {
        if (obj == null)
        {
            return null;
        }

        return new MobileDeviceListItemDto
        {
            Identity = GetString(obj, "Identity"),
            Guid = GetNullableString(obj, "Guid"),
            DeviceId = GetString(obj, "DeviceId"),
            FriendlyName = GetNullableString(obj, "FriendlyName"),
            DeviceType = GetNullableString(obj, "DeviceType"),
            DeviceModel = GetNullableString(obj, "DeviceModel"),
            DeviceUserAgent = GetNullableString(obj, "DeviceUserAgent"),
            DeviceOS = GetNullableString(obj, "DeviceOS"),
            ClientType = GetNullableString(obj, "ClientType"),
            UserDisplayName = GetNullableString(obj, "UserDisplayName"),
            UserPrincipalName = GetNullableString(obj, "UserPrincipalName"),
            MailboxIdentity = GetString(obj, "MailboxIdentity"),
            MailboxDisplayName = GetString(obj, "MailboxDisplayName"),
            CurrentMailboxPolicy = GetNullableString(obj, "CurrentMailboxPolicy"),
            DeviceAccessState = GetString(obj, "DeviceAccessState"),
            DeviceAccessStateReason = GetNullableString(obj, "DeviceAccessStateReason"),
            FirstSyncTime = GetNullableDateTime(obj, "FirstSyncTime"),
            LastSuccessSync = GetNullableDateTime(obj, "LastSuccessSync"),
            LastPolicyUpdateTime = GetNullableDateTime(obj, "LastPolicyUpdateTime"),
            Status = GetNullableString(obj, "Status")
        };
    }

    private static void ApplyMobileDeviceStatistics(MobileDeviceListItemDto device, PSObject? stats)
    {
        if (stats == null)
        {
            return;
        }

        device.FriendlyName ??= GetNullableString(stats, "FriendlyName");
        device.DeviceUserAgent ??= GetNullableString(stats, "DeviceUserAgent");
        device.DeviceOS ??= GetNullableString(stats, "DeviceOS");
        device.UserDisplayName ??= GetNullableString(stats, "UserDisplayName");
        device.UserPrincipalName = GetNullableString(stats, "UserPrincipalName") ?? device.UserPrincipalName;
        device.MailboxIdentity = ResolveDisplayedMailboxIdentity(device.MailboxIdentity, device.UserPrincipalName);
        device.MailboxDisplayName = ResolveDisplayedMailboxLabel(device.MailboxIdentity, device.UserDisplayName, device.UserPrincipalName);
        device.FirstSyncTime = GetNullableDateTime(stats, "FirstSyncTime") ?? device.FirstSyncTime;
        device.LastSuccessSync = GetNullableDateTime(stats, "LastSuccessSync") ?? device.LastSuccessSync;
        device.LastPolicyUpdateTime = GetNullableDateTime(stats, "LastPolicyUpdateTime") ?? device.LastPolicyUpdateTime;
        device.Status = GetNullableString(stats, "Status") ?? device.Status;
    }

    private static void ApplyMobileDeviceCasMailbox(MobileDeviceListItemDto device, PSObject? cas)
    {
        if (cas == null)
        {
            return;
        }

        device.CurrentMailboxPolicy = GetNullableString(cas, "CurrentMailboxPolicy") ?? device.CurrentMailboxPolicy;
        device.UserPrincipalName = GetNullableString(cas, "UserPrincipalName") ?? device.UserPrincipalName;
        device.MailboxIdentity = ResolveDisplayedMailboxIdentity(
            GetNullableString(cas, "MailboxIdentity") ?? device.MailboxIdentity,
            device.UserPrincipalName);
        device.MailboxDisplayName = ResolveDisplayedMailboxLabel(
            device.MailboxIdentity,
            GetNullableString(cas, "MailboxDisplayName") ?? device.UserDisplayName,
            device.UserPrincipalName);
    }

    internal static string ResolveDisplayedMailboxIdentity(string mailboxIdentity, string? userPrincipalName)
    {
        var normalizedMailboxIdentity = string.IsNullOrWhiteSpace(mailboxIdentity)
            ? string.Empty
            : mailboxIdentity.Trim();

        if (!string.IsNullOrWhiteSpace(userPrincipalName))
        {
            var normalizedUserPrincipalName = userPrincipalName.Trim();
            if (normalizedUserPrincipalName.Contains('@'))
            {
                return normalizedUserPrincipalName;
            }
        }

        return normalizedMailboxIdentity;
    }

    internal static string ResolveDisplayedMailboxLabel(string mailboxIdentity, string? userDisplayName, string? userPrincipalName)
    {
        if (!string.IsNullOrWhiteSpace(userDisplayName))
        {
            return userDisplayName.Trim();
        }

        if (!string.IsNullOrWhiteSpace(userPrincipalName))
        {
            var normalizedUserPrincipalName = userPrincipalName.Trim();
            if (normalizedUserPrincipalName.Contains('@'))
            {
                return normalizedUserPrincipalName;
            }
        }

        return string.IsNullOrWhiteSpace(mailboxIdentity)
            ? string.Empty
            : mailboxIdentity.Trim();
    }
}

