using System.Collections;
using System.Linq;
using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

public class ExoPermissionCommands
{
    private static readonly HashSet<string> AllowedSortProperties = new(StringComparer.OrdinalIgnoreCase)
    {
        "Name",
        "DisplayName",
        "WhenCreated",
        "WhenChanged"
    };

    private readonly PowerShellEngine _engine;

    public ExoPermissionCommands(PowerShellEngine engine)
    {
        _engine = engine;
    }

    public async Task<GetRoleGroupsResponse> GetRoleGroupsAsync(
        GetRoleGroupsRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var response = new GetRoleGroupsResponse
        {
            Skip = request.Skip,
            PageSize = request.PageSize,
            SearchQuery = request.SearchQuery
        };

        var escapedSearch = EscapePowerShellString(request.SearchQuery);
        var sortProperty = NormalizeSortProperty(request.SortBy);
        var sortDirection = request.SortDescending ? "-Descending" : string.Empty;

        var script = $@"
$allGroups = Get-RoleGroup | ForEach-Object {{
    $group = $_
    $members = @($group.Members | ForEach-Object {{ $_.ToString() }})
    $roles = @($group.Roles | ForEach-Object {{ $_.ToString() }})

    [pscustomobject]@{{
        Identity = $group.Identity.ToString()
        Name = $group.Name
        DisplayName = if ($group.DisplayName) {{ $group.DisplayName }} else {{ $group.Name }}
        Description = $group.Description
        MemberCount = $members.Count
        RoleCount = $roles.Count
        ManagedBy = @($group.ManagedBy | ForEach-Object {{ $_.ToString() }})
        WhenCreated = $group.WhenCreated
        WhenChanged = $group.WhenChanged
    }}
}}

if ('{escapedSearch}') {{
    $allGroups = $allGroups | Where-Object {{
        $_.Name -like '*{escapedSearch}*' -or
        $_.DisplayName -like '*{escapedSearch}*' -or
        $_.Description -like '*{escapedSearch}*'
    }}
}}

$allGroups = $allGroups | Sort-Object {sortProperty} {sortDirection}
$totalCount = @($allGroups).Count
$pagedGroups = $allGroups | Select-Object -Skip {request.Skip} -First {request.PageSize}

@{{
    TotalCount = $totalCount
    RoleGroups = @($pagedGroups | ForEach-Object {{
        @{{
            Identity = $_.Identity
            Name = $_.Name
            DisplayName = $_.DisplayName
            Description = $_.Description
            MemberCount = $_.MemberCount
            RoleCount = $_.RoleCount
            ManagedBy = @($_.ManagedBy | ForEach-Object {{ $_.ToString() }})
            WhenCreated = $_.WhenCreated
            WhenChanged = $_.WhenChanged
        }}
    }})
}}
";

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to fetch role groups: {result.ErrorMessage}");
        }

        if (result.Output.FirstOrDefault()?.BaseObject is Hashtable hash)
        {
            response.TotalCount = Convert.ToInt32(hash["TotalCount"] ?? 0);
            foreach (var item in ConvertToHashtableList(hash["RoleGroups"]))
            {
                response.RoleGroups.Add(new RoleGroupListItemDto
                {
                    Identity = item["Identity"]?.ToString() ?? string.Empty,
                    Name = item["Name"]?.ToString() ?? string.Empty,
                    DisplayName = item["DisplayName"]?.ToString() ?? string.Empty,
                    Description = item["Description"]?.ToString(),
                    MemberCount = Convert.ToInt32(item["MemberCount"] ?? 0),
                    RoleCount = Convert.ToInt32(item["RoleCount"] ?? 0),
                    ManagedBy = ConvertToStringList(item["ManagedBy"]),
                    WhenCreated = item["WhenCreated"] as DateTime?,
                    WhenChanged = item["WhenChanged"] as DateTime?
                });
            }
        }

        response.HasMore = request.Skip + response.RoleGroups.Count < response.TotalCount;
        return response;
    }

    public async Task<RoleGroupDetailsDto> GetRoleGroupDetailsAsync(
        GetRoleGroupDetailsRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var script = BuildGetRoleGroupDetailsScript(request);
        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success || result.Output.FirstOrDefault()?.BaseObject is not Hashtable hash)
        {
            throw new InvalidOperationException($"Failed to fetch role group details: {result.ErrorMessage}");
        }

        return new RoleGroupDetailsDto
        {
            Identity = hash["Identity"]?.ToString() ?? string.Empty,
            Name = hash["Name"]?.ToString() ?? string.Empty,
            DisplayName = hash["DisplayName"]?.ToString() ?? string.Empty,
            Description = hash["Description"]?.ToString(),
            Roles = ConvertToStringList(hash["Roles"]),
            Members = ConvertToRoleGroupMemberList(hash["Members"]),
            ManagedBy = ConvertToStringList(hash["ManagedBy"]),
            MemberCount = Convert.ToInt32(hash["MemberCount"] ?? 0),
            RoleCount = Convert.ToInt32(hash["RoleCount"] ?? 0),
            CustomRecipientWriteScope = hash["CustomRecipientWriteScope"]?.ToString(),
            CustomConfigWriteScope = hash["CustomConfigWriteScope"]?.ToString(),
            RecipientReadScope = hash["RecipientReadScope"]?.ToString(),
            RecipientWriteScope = hash["RecipientWriteScope"]?.ToString(),
            WhenCreated = hash["WhenCreated"] as DateTime?,
            WhenChanged = hash["WhenChanged"] as DateTime?
        };
    }

    internal static string BuildGetRoleGroupDetailsScript(GetRoleGroupDetailsRequest request)
    {
        var escapedIdentity = EscapePowerShellString(request.Identity);

        return $@"
function Test-RoleGroupMatchValue([object]$value, [string]$expected) {{
    if ($null -eq $value) {{
        return $false
    }}

    $candidate = $value.ToString().Trim()
    return -not [string]::IsNullOrWhiteSpace($candidate) -and
        $candidate.Equals($expected, [System.StringComparison]::OrdinalIgnoreCase)
}}

function Resolve-RoleGroup([string]$identity) {{
    if ([string]::IsNullOrWhiteSpace($identity)) {{
        throw 'Role group identity is required.'
    }}

    $normalizedIdentity = $identity.Trim()
    $allGroups = @(Get-RoleGroup)

    foreach ($propertyNames in @(
        @('Identity'),
        @('Guid'),
        @('DistinguishedName'),
        @('Name', 'DisplayName')
    )) {{
        $exactMatches = @($allGroups | Where-Object {{
            foreach ($propertyName in $propertyNames) {{
                $property = $_.PSObject.Properties[$propertyName]
                if ($null -ne $property -and (Test-RoleGroupMatchValue $property.Value $normalizedIdentity)) {{
                    return $true
                }}
            }}

            return $false
        }})

        if (@($exactMatches).Count -eq 1) {{
            return $exactMatches[0]
        }}

        if (@($exactMatches).Count -gt 1) {{
            $matchedProperties = $propertyNames -join ', '
            throw ""Multiple role groups matched identity '$normalizedIdentity' via $matchedProperties. Use a more specific identifier.""
        }}
    }}

    return Get-RoleGroup -Identity $normalizedIdentity -ErrorAction Stop
}}

function Get-ResolvedRoleGroupMembers([object]$group, [string]$requestedIdentity) {{
    $identityCandidates = @(
        if ($null -ne $group.Guid) {{ $group.Guid.ToString() }} else {{ $null }}
        if ($null -ne $group.Identity) {{ $group.Identity.ToString() }} else {{ $null }}
        if ($null -ne $group.Name) {{ $group.Name.ToString() }} else {{ $null }}
        $requestedIdentity
    ) | Where-Object {{ -not [string]::IsNullOrWhiteSpace($_) }} | Select-Object -Unique

    foreach ($candidate in $identityCandidates) {{
        try {{
            return @(Get-RoleGroupMember -Identity $candidate -ErrorAction Stop)
        }}
        catch {{
        }}
    }}

    throw ""Failed to fetch role group members for '$requestedIdentity'.""
}}

function Test-IsRoleGroupMemberTechnicalValue([string]$text) {{
    if ([string]::IsNullOrWhiteSpace($text)) {{
        return $true
    }}

    $candidate = $text.Trim()
    $guid = [Guid]::Empty
    if ([Guid]::TryParse($candidate, [ref]$guid)) {{
        return $true
    }}

    return $candidate -match '^S-\d-\d+-.+'
}}

function Get-RoleGroupMemberPropertyValue([object]$member, [string[]]$propertyNames, [switch]$SkipTechnicalValue) {{
    foreach ($propertyName in $propertyNames) {{
        $property = $member.PSObject.Properties[$propertyName]
        if ($null -eq $property -or $null -eq $property.Value) {{
            continue
        }}

        $text = $property.Value.ToString().Trim()
        if ([string]::IsNullOrWhiteSpace($text)) {{
            continue
        }}

        if ($SkipTechnicalValue -and (Test-IsRoleGroupMemberTechnicalValue $text)) {{
            continue
        }}

        return $text
    }}

    return $null
}}

function Resolve-RoleGroupMemberDisplayName([object]$member) {{
    $displayName = Get-RoleGroupMemberPropertyValue $member @('DisplayName', 'PrimarySmtpAddress', 'WindowsLiveID', 'UserPrincipalName', 'Alias', 'Name') -SkipTechnicalValue
    if (-not [string]::IsNullOrWhiteSpace($displayName)) {{
        return $displayName
    }}

    $identityText = if ($null -ne $member.Identity) {{ $member.Identity.ToString().Trim() }} else {{ $null }}
    if (-not [string]::IsNullOrWhiteSpace($identityText)) {{
        foreach ($resolver in @(
            {{ param($identity) Get-Recipient -Identity $identity -ErrorAction Stop }},
            {{ param($identity) Get-User -Identity $identity -ErrorAction Stop }}
        )) {{
            try {{
                $resolvedMember = & $resolver $identityText
                if ($null -eq $resolvedMember) {{
                    continue
                }}

                $resolvedDisplayName = Get-RoleGroupMemberPropertyValue $resolvedMember @('DisplayName', 'PrimarySmtpAddress', 'WindowsLiveID', 'UserPrincipalName', 'Alias', 'Name') -SkipTechnicalValue
                if (-not [string]::IsNullOrWhiteSpace($resolvedDisplayName)) {{
                    return $resolvedDisplayName
                }}
            }}
            catch {{
            }}
        }}
    }}

    $fallbackDisplayName = Get-RoleGroupMemberPropertyValue $member @('DisplayName', 'PrimarySmtpAddress', 'WindowsLiveID', 'UserPrincipalName', 'Alias', 'Name')
    if (-not [string]::IsNullOrWhiteSpace($fallbackDisplayName)) {{
        return $fallbackDisplayName
    }}

    if (-not [string]::IsNullOrWhiteSpace($identityText)) {{
        return $identityText
    }}

    return '(unknown member)'
}}

$group = Resolve-RoleGroup '{escapedIdentity}'
$members = Get-ResolvedRoleGroupMembers $group '{escapedIdentity}' | ForEach-Object {{
    @{{
        Identity = if ($_.Identity) {{ $_.Identity.ToString() }} else {{ Resolve-RoleGroupMemberDisplayName $_ }}
        DisplayName = Resolve-RoleGroupMemberDisplayName $_
    }}
}}

@{{
    Identity = $group.Identity.ToString()
    Name = $group.Name
    DisplayName = if ($group.DisplayName) {{ $group.DisplayName }} else {{ $group.Name }}
    Description = $group.Description
    Roles = @($group.Roles | ForEach-Object {{ $_.ToString() }})
    Members = @($members | ForEach-Object {{
        @{{
            Identity = $_.Identity
            DisplayName = $_.DisplayName
        }}
    }})
    ManagedBy = @($group.ManagedBy | ForEach-Object {{ $_.ToString() }})
    MemberCount = @($members).Count
    RoleCount = @($group.Roles).Count
    CustomRecipientWriteScope = if ($group.CustomRecipientWriteScope) {{ $group.CustomRecipientWriteScope.ToString() }} else {{ $null }}
    CustomConfigWriteScope = if ($group.CustomConfigWriteScope) {{ $group.CustomConfigWriteScope.ToString() }} else {{ $null }}
    RecipientReadScope = if ($group.RecipientReadScope) {{ $group.RecipientReadScope.ToString() }} else {{ $null }}
    RecipientWriteScope = if ($group.RecipientWriteScope) {{ $group.RecipientWriteScope.ToString() }} else {{ $null }}
    WhenCreated = $group.WhenCreated
    WhenChanged = $group.WhenChanged
}}
";
    }

    public async Task UpsertRoleGroupAsync(
        UpsertRoleGroupRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.Name))
        {
            throw new InvalidOperationException("Name is required.");
        }

        var escapedName = EscapePowerShellString(request.Name);
        var escapedIdentity = EscapePowerShellString(request.Identity);
        var escapedDescription = EscapePowerShellString(request.Description);
        var escapedCopySource = EscapePowerShellString(request.CopyFromRoleGroup);
        var rolesParameter = ExoRequestSanitizer.FormatStringArrayParameter(request.Roles);
        var membersParameter = ExoRequestSanitizer.FormatStringArrayParameter(request.Members);

        var script = string.IsNullOrWhiteSpace(request.Identity)
            ? $@"
$roles = {rolesParameter}
$members = {membersParameter}

if ('{escapedCopySource}') {{
    $source = Get-RoleGroup -Identity '{escapedCopySource}'
    if ($roles -eq $null) {{
        $roles = @($source.Roles | ForEach-Object {{ $_.ToString() }})
    }}

    if ($members -eq $null) {{
        $members = Get-RoleGroupMember -Identity '{escapedCopySource}' | ForEach-Object {{
            if ($_.Name) {{ $_.Name }} else {{ $_.Identity.ToString() }}
        }}
    }}
}}

if ($roles -eq $null -or @($roles).Count -eq 0) {{
    throw 'At least one management role is required to create a role group.'
}}

$newParams = @{{
    Name = '{escapedName}'
    Roles = $roles
}}

if ('{escapedDescription}') {{
    $newParams['Description'] = '{escapedDescription}'
}}

if ($members -ne $null -and @($members).Count -gt 0) {{
    $newParams['Members'] = $members
}}

New-RoleGroup @newParams | Out-Null
"
            : $@"
$setParams = @{{
    Identity = '{escapedIdentity}'
}}

if ('{escapedDescription}') {{
    $setParams['Description'] = '{escapedDescription}'
}}

if ({(request.Roles.Count > 0 ? "$true" : "$false")}) {{
    $setParams['Roles'] = {rolesParameter}
}}

Set-RoleGroup @setParams | Out-Null
";

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to save role group: {result.ErrorMessage}");
        }
    }

    public async Task ModifyRoleGroupMemberAsync(
        ModifyRoleGroupMemberRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.Identity) || string.IsNullOrWhiteSpace(request.Member))
        {
            throw new InvalidOperationException("Identity and member are required.");
        }

        var escapedIdentity = EscapePowerShellString(request.Identity);
        var escapedMember = EscapePowerShellString(request.Member);
        var script = request.Action == RoleGroupMemberAction.Add
            ? $"Add-RoleGroupMember -Identity '{escapedIdentity}' -Member '{escapedMember}' | Out-Null"
            : $"Remove-RoleGroupMember -Identity '{escapedIdentity}' -Member '{escapedMember}' -Confirm:$false | Out-Null";

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to update role group membership: {result.ErrorMessage}");
        }
    }

    private static string NormalizeSortProperty(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "DisplayName";
        }

        var normalized = value.Trim();
        return AllowedSortProperties.Contains(normalized) ? normalized : "DisplayName";
    }

    private static string EscapePowerShellString(string? value) => value?.Replace("'", "''") ?? string.Empty;

    private static IEnumerable<Hashtable> ConvertToHashtableList(object? value)
    {
        if (value is Hashtable single)
        {
            return new[] { single };
        }

        if (value is object[] array)
        {
            return array.OfType<Hashtable>();
        }

        return value is IEnumerable enumerable
            ? enumerable.OfType<Hashtable>()
            : Enumerable.Empty<Hashtable>();
    }

    private static List<RoleGroupMemberDto> ConvertToRoleGroupMemberList(object? value)
    {
        return ConvertToHashtableList(value)
            .Select(item =>
            {
                var identity = item["Identity"]?.ToString() ?? string.Empty;
                var displayName = item["DisplayName"]?.ToString() ?? string.Empty;
                return new RoleGroupMemberDto
                {
                    Identity = identity,
                    DisplayName = string.IsNullOrWhiteSpace(displayName) ? identity : displayName
                };
            })
            .Where(member => !string.IsNullOrWhiteSpace(member.Identity) || !string.IsNullOrWhiteSpace(member.DisplayName))
            .ToList();
    }

    private static List<string> ConvertToStringList(object? value)
    {
        if (value == null)
        {
            return new List<string>();
        }

        if (value is string single)
        {
            return string.IsNullOrWhiteSpace(single) ? new List<string>() : new List<string> { single };
        }

        if (value is IEnumerable enumerable)
        {
            var list = new List<string>();
            foreach (var item in enumerable)
            {
                var text = item?.ToString();
                if (!string.IsNullOrWhiteSpace(text))
                {
                    list.Add(text);
                }
            }

            return list.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(item => item, StringComparer.OrdinalIgnoreCase).ToList();
        }

        var fallback = value.ToString();
        return string.IsNullOrWhiteSpace(fallback) ? new List<string>() : new List<string> { fallback };
    }
}

