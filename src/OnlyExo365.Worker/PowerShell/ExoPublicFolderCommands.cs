using System.Collections;
using System.Management.Automation;
using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

internal sealed class ExoPublicFolderCommands : ExoCommandModuleBase
{
    public ExoPublicFolderCommands(PowerShellEngine engine)
        : base(engine)
    {
    }

    public async Task<GetPublicFoldersResponse> GetPublicFoldersAsync(
        GetPublicFoldersRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var sortProperty = NormalizePublicFolderSortProperty(request.SortBy);
        var mailEnabledFilter = request.MailEnabledOnly switch
        {
            true => "$true",
            false => "$false",
            _ => "$null"
        };

        var response = new GetPublicFoldersResponse
        {
            Skip = request.Skip,
            PageSize = request.PageSize,
            SearchQuery = request.SearchQuery
        };

        if (!string.IsNullOrWhiteSpace(request.SortBy) &&
            !string.Equals(sortProperty, request.SortBy.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            onLog?.Invoke("Warning", $"Unsupported SortBy ignored: {request.SortBy}");
        }

        var escapedSearch = EscapePs(request.SearchQuery);
        var sortDirection = request.SortDescending ? "-Descending" : string.Empty;

        var script = $@"
$mailFoldersByIdentity = @{{}}
try {{
    Get-MailPublicFolder -ResultSize Unlimited -ErrorAction Stop | ForEach-Object {{
        $mailFoldersByIdentity[$_.Identity.ToString()] = $_
    }}
}}
catch {{
    Write-Warning ""Get-MailPublicFolder failed: $($_.Exception.Message)""
}}

$items = Get-PublicFolder -Identity '\' -Recurse -ErrorAction Stop |
    Where-Object {{ $_.Identity.ToString() -ne '\' }} |
    ForEach-Object {{
        $folderIdentity = $_.Identity.ToString()
        $parentPath = Split-Path -Path $folderIdentity -Parent
        if ([string]::IsNullOrWhiteSpace($parentPath)) {{
            $parentPath = '\'
        }}

        $mailFolder = $null
        if ($mailFoldersByIdentity.ContainsKey($folderIdentity)) {{
            $mailFolder = $mailFoldersByIdentity[$folderIdentity]
        }}

        [PSCustomObject]@{{
            Identity = $folderIdentity
            Name = $_.Name
            ParentPath = $parentPath
            MailEnabled = ($null -ne $mailFolder)
            Alias = if ($mailFolder -and $mailFolder.Alias) {{ $mailFolder.Alias.ToString() }} else {{ $null }}
            PrimarySmtpAddress = if ($mailFolder -and $mailFolder.PrimarySmtpAddress) {{ $mailFolder.PrimarySmtpAddress.ToString() }} else {{ $null }}
            HiddenFromAddressListsEnabled = if ($mailFolder) {{ [bool]$mailFolder.HiddenFromAddressListsEnabled }} else {{ $false }}
            HasSubFolders = [bool]$_.HasSubFolders
        }}
    }}

$searchQuery = '{escapedSearch}'
if (-not [string]::IsNullOrWhiteSpace($searchQuery)) {{
    $items = $items | Where-Object {{
        $_.Identity -like ""*$searchQuery*"" -or
        $_.Name -like ""*$searchQuery*"" -or
        $_.Alias -like ""*$searchQuery*"" -or
        $_.PrimarySmtpAddress -like ""*$searchQuery*""
    }}
}}

$mailEnabledOnly = {mailEnabledFilter}
if ($null -ne $mailEnabledOnly) {{
    $items = $items | Where-Object {{ $_.MailEnabled -eq $mailEnabledOnly }}
}}

$items = $items | Sort-Object {sortProperty} {sortDirection}
$totalCount = @($items).Count
$pagedItems = $items | Select-Object -Skip {request.Skip} -First {request.PageSize}

@{{
    TotalCount = $totalCount
    Folders = @($pagedItems)
}}";

        onLog?.Invoke("Verbose", $"Fetching public folders (skip={request.Skip}, pageSize={request.PageSize})...");

        var result = await Engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (result.Success && result.Output.Any() && result.Output.First().BaseObject is Hashtable hash)
        {
            response.TotalCount = Convert.ToInt32(hash["TotalCount"] ?? 0);

            if (hash["Folders"] is object[] folders)
            {
                foreach (var folderObject in folders)
                {
                    if (folderObject is not PSObject folderPs)
                    {
                        continue;
                    }

                    response.Folders.Add(new PublicFolderListItemDto
                    {
                        Identity = GetString(folderPs, "Identity"),
                        Name = GetString(folderPs, "Name"),
                        ParentPath = GetString(folderPs, "ParentPath"),
                        MailEnabled = GetBool(folderPs, "MailEnabled"),
                        Alias = GetNullableString(folderPs, "Alias"),
                        PrimarySmtpAddress = GetNullableString(folderPs, "PrimarySmtpAddress"),
                        HiddenFromAddressListsEnabled = GetBool(folderPs, "HiddenFromAddressListsEnabled"),
                        HasSubFolders = GetBool(folderPs, "HasSubFolders")
                    });
                }
            }

            response.HasMore = (request.Skip + response.Folders.Count) < response.TotalCount;
        }

        onLog?.Invoke("Information", $"Retrieved {response.Folders.Count} public folders (total: {response.TotalCount})");

        return response;
    }

    public async Task<PublicFolderDetailsDto> GetPublicFolderDetailsAsync(
        GetPublicFolderDetailsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var escapedIdentity = EscapePs(request.Identity);

        var script = $@"
$folder = Get-PublicFolder -Identity '{escapedIdentity}' -ErrorAction Stop
$mailFolder = $null
try {{
    $mailFolder = Get-MailPublicFolder -Identity '{escapedIdentity}' -ErrorAction Stop
}}
catch {{
}}

$stats = $null
try {{
    $stats = Get-PublicFolderStatistics -Identity '{escapedIdentity}' -ErrorAction Stop
}}
catch {{
}}

$permissions = @()
try {{
    $permissions = Get-PublicFolderClientPermission -Identity '{escapedIdentity}' -ErrorAction Stop | ForEach-Object {{
        [PSCustomObject]@{{
            User = $_.User.ToString()
            AccessRights = @($_.AccessRights | ForEach-Object {{ $_.ToString() }})
        }}
    }}
}}
catch {{
}}

$parentPath = Split-Path -Path $folder.Identity.ToString() -Parent
if ([string]::IsNullOrWhiteSpace($parentPath)) {{
    $parentPath = '\'
}}

[PSCustomObject]@{{
    Identity = $folder.Identity.ToString()
    Name = $folder.Name
    ParentPath = $parentPath
    MailEnabled = ($null -ne $mailFolder)
    Alias = if ($mailFolder -and $mailFolder.Alias) {{ $mailFolder.Alias.ToString() }} else {{ $null }}
    PrimarySmtpAddress = if ($mailFolder -and $mailFolder.PrimarySmtpAddress) {{ $mailFolder.PrimarySmtpAddress.ToString() }} else {{ $null }}
    HiddenFromAddressListsEnabled = if ($mailFolder) {{ [bool]$mailFolder.HiddenFromAddressListsEnabled }} else {{ $false }}
    HasSubFolders = [bool]$folder.HasSubFolders
    ItemCount = if ($stats -and $null -ne $stats.ItemCount) {{ [int]$stats.ItemCount }} else {{ $null }}
    TotalItemSize = if ($stats -and $stats.TotalItemSize) {{ $stats.TotalItemSize.ToString() }} else {{ $null }}
    Permissions = @($permissions)
}}";

        onLog?.Invoke("Verbose", $"Fetching public folder details for {request.Identity}...");

        var results = await RunScriptAsync(script, cancellationToken);
        if (results.Count == 0)
        {
            throw new InvalidOperationException($"Public folder not found: {request.Identity}");
        }

        var obj = results[0];
        var details = new PublicFolderDetailsDto
        {
            Identity = GetString(obj, "Identity"),
            Name = GetString(obj, "Name"),
            ParentPath = GetString(obj, "ParentPath"),
            MailEnabled = GetBool(obj, "MailEnabled"),
            Alias = GetNullableString(obj, "Alias"),
            PrimarySmtpAddress = GetNullableString(obj, "PrimarySmtpAddress"),
            HiddenFromAddressListsEnabled = GetBool(obj, "HiddenFromAddressListsEnabled"),
            HasSubFolders = GetBool(obj, "HasSubFolders"),
            ItemCount = GetNullableInt(obj, "ItemCount"),
            TotalItemSize = GetNullableString(obj, "TotalItemSize")
        };

        if (obj.Properties["Permissions"]?.Value is object[] permissions)
        {
            foreach (var permissionObject in permissions)
            {
                if (permissionObject is not PSObject permissionPs)
                {
                    continue;
                }

                details.Permissions.Add(new PublicFolderPermissionEntryDto
                {
                    User = GetString(permissionPs, "User"),
                    DisplayName = GetString(permissionPs, "User"),
                    AccessRights = ConvertToStringList(permissionPs.Properties["AccessRights"]?.Value)
                });
            }
        }

        return details;
    }

    public async Task<UpsertPublicFolderResponse> UpsertPublicFolderAsync(
        UpsertPublicFolderRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var script = BuildUpsertPublicFolderScript(request);

        onLog?.Invoke("Information", $"Saving public folder {request.Name}...");

        var results = await RunScriptAsync(script, cancellationToken);
        var obj = results.LastOrDefault();

        return new UpsertPublicFolderResponse
        {
            Identity = obj == null ? request.Identity ?? request.Name : GetString(obj, "Identity"),
            MailEnabled = obj != null && GetBool(obj, "MailEnabled"),
            PrimarySmtpAddress = obj == null ? request.PrimarySmtpAddress : GetNullableString(obj, "PrimarySmtpAddress")
        };
    }

    public async Task SetPublicFolderClientPermissionAsync(
        SetPublicFolderClientPermissionRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var result = await Engine.ExecuteAsync(
            BuildSetPublicFolderClientPermissionScript(request),
            onVerbose: onLog,
            cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException(
                $"Failed to {request.Action} public folder client permission on {request.Identity}: {result.ErrorMessage}");
        }
    }

    public async Task RemovePublicFolderAsync(
        RemovePublicFolderRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var result = await Engine.ExecuteAsync(
            BuildRemovePublicFolderScript(request),
            onVerbose: onLog,
            cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to remove public folder {request.Identity}: {result.ErrorMessage}");
        }
    }

    internal static string BuildUpsertPublicFolderScript(UpsertPublicFolderRequest request)
    {
        var identity = EscapePs(request.Identity);
        var name = EscapePs(request.Name);
        var parentPath = EscapePs(string.IsNullOrWhiteSpace(request.ParentPath) ? "\\" : request.ParentPath);
        var alias = EscapePs(request.Alias);
        var primarySmtpAddress = EscapePs(request.PrimarySmtpAddress);
        var hidden = ToPsBoolLiteral(request.HiddenFromAddressListsEnabled);
        var mailEnabled = ToPsBoolLiteral(request.MailEnabled);

        return $@"
$requestedIdentity = '{identity}'
$name = '{name}'
$parentPath = '{parentPath}'
$mailEnabled = {mailEnabled}
$alias = '{alias}'
$primarySmtpAddress = '{primarySmtpAddress}'
$hidden = {hidden}

{BuildPublicFolderPathHelpers()}

if ([string]::IsNullOrWhiteSpace($requestedIdentity)) {{
    New-PublicFolder -Name $name -Path $parentPath -ErrorAction Stop | Out-Null
    $targetIdentity = Join-PublicFolderPath $parentPath $name
}}
else {{
    $folder = Get-PublicFolder -Identity $requestedIdentity -ErrorAction Stop
    $targetIdentity = $requestedIdentity
    $currentParentPath = Split-Path -Path $requestedIdentity -Parent
    if ([string]::IsNullOrWhiteSpace($currentParentPath)) {{
        $currentParentPath = '\'
    }}

    if ($folder.Name -ne $name) {{
        Set-PublicFolder -Identity $requestedIdentity -Name $name -ErrorAction Stop
        $targetIdentity = Join-PublicFolderPath $currentParentPath $name
    }}

    if ($parentPath -ne $currentParentPath) {{
        Set-PublicFolder -Identity $targetIdentity -Path $parentPath -ErrorAction Stop
        $targetIdentity = Join-PublicFolderPath $parentPath $name
    }}
}}

$mailFolder = $null
try {{
    $mailFolder = Get-MailPublicFolder -Identity $targetIdentity -ErrorAction Stop
}}
catch {{
}}

if ($mailEnabled) {{
    if ($null -eq $mailFolder) {{
        Enable-MailPublicFolder -Identity $targetIdentity -Alias $alias -ErrorAction Stop | Out-Null
    }}

    $setParams = @{{
        Identity = $targetIdentity
        HiddenFromAddressListsEnabled = $hidden
        ErrorAction = 'Stop'
    }}

    if (-not [string]::IsNullOrWhiteSpace($alias)) {{
        $setParams['Alias'] = $alias
    }}

    if (-not [string]::IsNullOrWhiteSpace($primarySmtpAddress)) {{
        $setParams['PrimarySmtpAddress'] = $primarySmtpAddress
    }}

    Set-MailPublicFolder @setParams | Out-Null
}}
elseif ($mailFolder) {{
    Disable-MailPublicFolder -Identity $targetIdentity -Confirm:$false -ErrorAction Stop | Out-Null
}}

$finalMailFolder = $null
try {{
    $finalMailFolder = Get-MailPublicFolder -Identity $targetIdentity -ErrorAction Stop
}}
catch {{
}}

[PSCustomObject]@{{
    Identity = $targetIdentity
    MailEnabled = ($null -ne $finalMailFolder)
    PrimarySmtpAddress = if ($finalMailFolder -and $finalMailFolder.PrimarySmtpAddress) {{ $finalMailFolder.PrimarySmtpAddress.ToString() }} else {{ $null }}
}}";
    }

    internal static string BuildSetPublicFolderClientPermissionScript(SetPublicFolderClientPermissionRequest request)
    {
        var identity = EscapePs(request.Identity);
        var user = EscapePs(request.User);
        var action = request.Action.ToString();
        var accessRights = ToPsArrayLiteral(request.AccessRights);

        return $@"
$identity = '{identity}'
$user = '{user}'
$action = '{action}'
$accessRights = {accessRights}

switch ($action) {{
    'Add' {{
        Add-PublicFolderClientPermission -Identity $identity -User $user -AccessRights $accessRights -ErrorAction Stop | Out-Null
    }}
    'Modify' {{
        Remove-PublicFolderClientPermission -Identity $identity -User $user -Confirm:$false -ErrorAction Stop | Out-Null
        Add-PublicFolderClientPermission -Identity $identity -User $user -AccessRights $accessRights -ErrorAction Stop | Out-Null
    }}
    'Remove' {{
        Remove-PublicFolderClientPermission -Identity $identity -User $user -Confirm:$false -ErrorAction Stop | Out-Null
    }}
    default {{
        throw ""Unsupported public folder client permission action: $action""
    }}
}}";
    }

    internal static string BuildRemovePublicFolderScript(RemovePublicFolderRequest request)
    {
        var identity = EscapePs(request.Identity);
        var recursive = ToPsBoolLiteral(request.Recursive);

        return $@"
$identity = '{identity}'
$recursive = {recursive}

$mailFolder = $null
try {{
    $mailFolder = Get-MailPublicFolder -Identity $identity -ErrorAction Stop
}}
catch {{
}}

if ($mailFolder) {{
    Disable-MailPublicFolder -Identity $identity -Confirm:$false -ErrorAction Stop | Out-Null
}}

$removeParams = @{{
    Identity = $identity
    Confirm = $false
    ErrorAction = 'Stop'
}}

if ($recursive) {{
    $removeParams['Recurse'] = $true
}}

Remove-PublicFolder @removeParams | Out-Null";
    }

    internal static string NormalizePublicFolderSortProperty(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "Name";
        }

        return value.Trim() switch
        {
            "Name" => "Name",
            "Identity" => "Identity",
            "Path" => "Identity",
            "PrimarySmtpAddress" => "PrimarySmtpAddress",
            "MailEnabled" => "MailEnabled",
            _ => "Name"
        };
    }

    private static string BuildPublicFolderPathHelpers()
    {
        return @"
function Join-PublicFolderPath([string]$basePath, [string]$childName) {
    if ([string]::IsNullOrWhiteSpace($basePath) -or $basePath -eq '\') {
        return '\' + $childName
    }

    if ($basePath.EndsWith('\')) {
        return $basePath + $childName
    }

    return $basePath + '\' + $childName
}";
    }
}

