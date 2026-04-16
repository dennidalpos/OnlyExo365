using ExchangeAdmin.Contracts.Dtos;
using System.Linq;
using System.Management.Automation;

namespace ExchangeAdmin.Worker.PowerShell;

internal sealed partial class ExoMailboxReportingCommands
{
    public async Task<GetMailboxFolderPermissionsResponse> GetMailboxFolderPermissionsAsync(
        GetMailboxFolderPermissionsRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var results = await RunScriptAsync(BuildGetMailboxFolderPermissionsScript(request), cancellationToken);
        var response = new GetMailboxFolderPermissionsResponse
        {
            MailboxIdentity = request.MailboxIdentity,
            FolderPath = NormalizeMailboxFolderPath(request.FolderPath)
        };

        var obj = results.FirstOrDefault();
        if (obj == null)
        {
            return response;
        }

        response.ResolvedFolderIdentity = GetString(obj, "ResolvedFolderIdentity");
        if (obj.Properties["Permissions"]?.Value is not object[] permissions)
        {
            return response;
        }

        foreach (var permissionObject in permissions)
        {
            if (permissionObject is not PSObject permissionPs)
            {
                continue;
            }

            response.Permissions.Add(new MailboxFolderPermissionEntryDto
            {
                User = GetString(permissionPs, "User"),
                DisplayName = GetString(permissionPs, "DisplayName"),
                AccessRights = ConvertToStringList(permissionPs.Properties["AccessRights"]?.Value),
                IsInherited = GetBool(permissionPs, "IsInherited")
            });
        }

        onLog?.Invoke("Information", $"Retrieved {response.Permissions.Count} folder permissions for {response.ResolvedFolderIdentity}");

        return response;
    }

    public async Task SetMailboxFolderPermissionAsync(
        SetMailboxFolderPermissionRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var script = BuildSetMailboxFolderPermissionScript(request);
        var result = await Engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException(
                $"Failed to {request.Action} mailbox folder permission on {request.MailboxIdentity}:{request.FolderPath}: {result.ErrorMessage}");
        }
    }

    internal static string BuildGetMailboxFolderPermissionsScript(GetMailboxFolderPermissionsRequest request)
    {
        var mailboxIdentity = EscapePs(request.MailboxIdentity);
        var folderPath = EscapePs(NormalizeMailboxFolderPath(request.FolderPath));

        return $@"
$mailboxIdentity = '{mailboxIdentity}'
$folderPath = '{folderPath}'

{BuildMailboxFolderIdentityResolverScript()}

$resolvedFolderIdentity = Resolve-MailboxFolderIdentity -MailboxIdentity $mailboxIdentity -FolderPath $folderPath
$permissions = @(Get-MailboxFolderPermission -Identity $resolvedFolderIdentity -ErrorAction Stop | ForEach-Object {{
    $userValue = $_.User.ToString()
    [PSCustomObject]@{{
        User = $userValue
        DisplayName = $userValue
        AccessRights = @($_.AccessRights | ForEach-Object {{ $_.ToString() }})
        IsInherited = [bool]$_.IsInherited
    }}
}})

[PSCustomObject]@{{
    ResolvedFolderIdentity = $resolvedFolderIdentity
    Permissions = @($permissions)
}}";
    }

    internal static string BuildSetMailboxFolderPermissionScript(SetMailboxFolderPermissionRequest request)
    {
        var mailboxIdentity = EscapePs(request.MailboxIdentity);
        var folderPath = EscapePs(NormalizeMailboxFolderPath(request.FolderPath));
        var user = EscapePs(request.User);
        var action = request.Action.ToString();
        var accessRights = ToPsArrayLiteral(request.AccessRights);

        return $@"
$mailboxIdentity = '{mailboxIdentity}'
$folderPath = '{folderPath}'
$user = '{user}'
$action = '{action}'
$accessRights = {accessRights}

{BuildMailboxFolderIdentityResolverScript()}

$resolvedFolderIdentity = Resolve-MailboxFolderIdentity -MailboxIdentity $mailboxIdentity -FolderPath $folderPath

switch ($action) {{
    'Add' {{
        Add-MailboxFolderPermission -Identity $resolvedFolderIdentity -User $user -AccessRights $accessRights -Confirm:$false -ErrorAction Stop | Out-Null
    }}
    'Modify' {{
        Set-MailboxFolderPermission -Identity $resolvedFolderIdentity -User $user -AccessRights $accessRights -Confirm:$false -ErrorAction Stop | Out-Null
    }}
    'Remove' {{
        Remove-MailboxFolderPermission -Identity $resolvedFolderIdentity -User $user -Confirm:$false -ErrorAction Stop | Out-Null
    }}
    default {{
        throw ""Unsupported mailbox folder permission action: $action""
    }}
}}";
    }

    private static string NormalizeMailboxFolderPath(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "Calendar";
        }

        return value
            .Trim()
            .Replace('/', '\\')
            .TrimStart('\\');
    }

    private static string BuildMailboxFolderIdentityResolverScript()
    {
        return @"
function Normalize-MailboxFolderSegment {
    param(
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }

    $normalized = $Value.Trim().Normalize([System.Text.NormalizationForm]::FormD)
    $builder = New-Object System.Text.StringBuilder
    foreach ($character in $normalized.ToCharArray()) {
        $unicodeCategory = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($character)
        if ($unicodeCategory -eq [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            continue
        }

        if ([char]::IsWhiteSpace($character) -or $character -eq '-' -or $character -eq '_') {
            continue
        }

        [void]$builder.Append([char]::ToLowerInvariant($character))
    }

    return $builder.ToString()
}

function Resolve-MailboxFolderIdentity {
    param(
        [string]$MailboxIdentity,
        [string]$FolderPath
    )

    $normalizedPath = ($FolderPath -replace '/', '\').Trim().TrimStart('\')
    if ([string]::IsNullOrWhiteSpace($normalizedPath)) {
        $normalizedPath = 'Calendar'
    }

    if ($normalizedPath -like '*:\*') {
        return $normalizedPath
    }

    $segments = $normalizedPath.Split('\')
    $rootSegment = if ($segments.Count -gt 0) { $segments[0] } else { $normalizedPath }
    $remainder = if ($segments.Count -gt 1) { '\' + (($segments | Select-Object -Skip 1) -join '\') } else { '' }

    $scopeMap = @{
        'calendar' = 'Calendar'
        'calendario' = 'Calendar'
        'contacts' = 'Contacts'
        'contatti' = 'Contacts'
        'tasks' = 'Tasks'
        'attivita' = 'Tasks'
        'notes' = 'Notes'
        'journal' = 'Journal'
        'diario' = 'Journal'
        'inbox' = 'Inbox'
        'postainarrivo' = 'Inbox'
        'sentitems' = 'SentItems'
        'postainviata' = 'SentItems'
        'deleteditems' = 'DeletedItems'
        'postaeliminata' = 'DeletedItems'
        'drafts' = 'Drafts'
        'bozze' = 'Drafts'
    }

    $normalizedRoot = Normalize-MailboxFolderSegment $rootSegment
    $scope = $scopeMap[$normalizedRoot]
    if ($scope) {
        try {
            $scopeEntries = @(Get-MailboxFolderStatistics -Identity $MailboxIdentity -FolderScope $scope -ErrorAction Stop)
            if ($scopeEntries.Count -gt 0) {
                $scopeEntry = @($scopeEntries |
                    Where-Object { $_.FolderType -and $_.FolderType.ToString() -eq $scope } |
                    Sort-Object @{ Expression = {
                            $folderPath = if ($_.FolderPath) { $_.FolderPath.ToString().Trim('/') } else { '' }
                            if ([string]::IsNullOrWhiteSpace($folderPath)) { [int]::MaxValue } else { $folderPath.Length }
                        } }) | Select-Object -First 1

                if (-not $scopeEntry) {
                    $scopeEntry = @($scopeEntries |
                        Where-Object { $_.FolderPath } |
                        Sort-Object @{ Expression = { $_.FolderPath.ToString().Trim('/').Length } }) | Select-Object -First 1
                }

                if ($scopeEntry -and $scopeEntry.FolderPath) {
                    $resolvedRootPath = $scopeEntry.FolderPath.ToString().Trim('/')
                    if (-not [string]::IsNullOrWhiteSpace($resolvedRootPath)) {
                        $resolvedRoot = ""${MailboxIdentity}:\$resolvedRootPath""
                        if ([string]::IsNullOrWhiteSpace($remainder)) {
                            return $resolvedRoot
                        }

                        return $resolvedRoot + $remainder
                    }
                }

                if ($scopeEntry -and $scopeEntry.Identity) {
                    $resolvedRoot = $scopeEntry.Identity.ToString()
                    if ([string]::IsNullOrWhiteSpace($remainder)) {
                        return $resolvedRoot
                    }

                    return $resolvedRoot + $remainder
                }
            }
        }
        catch {
        }
    }

    return ""${MailboxIdentity}:\$normalizedPath""
}";
    }
}
