using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

internal sealed class ExoOfflineAddressBookCommands : ExoCommandModuleBase
{
    private const string OfflineAddressBookCmdletUnavailableCode = "OfflineAddressBookCmdletUnavailable";

    public ExoOfflineAddressBookCommands(PowerShellEngine engine)
        : base(engine)
    {
    }

    public async Task<GetOfflineAddressBooksResponse> GetOfflineAddressBooksAsync(CancellationToken cancellationToken = default)
    {
        var script = BuildGetOfflineAddressBooksScript();
        var result = await Engine.ExecuteAsync(script, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException(result.ErrorMessage ?? "Offline address book query failed");
        }

        return MapOfflineAddressBooksResponse(result.Output, result.Warning);
    }

    internal static string BuildGetOfflineAddressBooksScript()
    {
        return @"
if (-not (Get-Command -Name 'Get-OfflineAddressBook' -ErrorAction SilentlyContinue)) {
    $warningPayload = @{
        Code = '__CODE__'
        Scope = 'MailFlow.OfflineAddressBooks'
        Message = 'Offline Address Books are not supported in the current Exchange session: Get-OfflineAddressBook is not available.'
        IsPartialData = $true
    } | ConvertTo-Json -Compress -Depth 3
    Write-Warning ('__WARN__' + $warningPayload)
    return
}

Get-OfflineAddressBook -ErrorAction Stop |
    Sort-Object Name |
    ForEach-Object {
        [PSCustomObject]@{
            Identity = $_.Identity.ToString()
            Name = $_.Name
            AddressLists = @($_.AddressLists | ForEach-Object { $_.ToString() })
            DiffRetentionPeriod = if ($null -ne $_.DiffRetentionPeriod) { [int]$_.DiffRetentionPeriod } else { $null }
            IsDefault = if ($null -ne $_.IsDefault) { [bool]$_.IsDefault } elseif ($null -ne $_.Default) { [bool]$_.Default } else { $false }
        }
    }"
            .Replace("__CODE__", OfflineAddressBookCmdletUnavailableCode, StringComparison.Ordinal)
            .Replace("__WARN__", StructuredWarningPrefix, StringComparison.Ordinal);
    }

    internal static GetOfflineAddressBooksResponse MapOfflineAddressBooksResponse(
        IEnumerable<System.Management.Automation.PSObject> results,
        IEnumerable<string>? warnings)
    {
        var warningDetails = ParseStructuredWarnings(warnings);
        var items = new List<OfflineAddressBookDto>();

        foreach (var obj in results)
        {
            items.Add(new OfflineAddressBookDto
            {
                Identity = GetString(obj, "Identity"),
                Name = GetString(obj, "Name"),
                AddressLists = ConvertToStringList(obj.Properties["AddressLists"]?.Value),
                DiffRetentionPeriod = GetNullableInt(obj, "DiffRetentionPeriod"),
                IsDefault = GetBool(obj, "IsDefault")
            });
        }

        return new GetOfflineAddressBooksResponse
        {
            OfflineAddressBooks = items,
            Warnings = ExtractWarningMessages(warningDetails),
            WarningDetails = warningDetails,
            HasPartialData = warningDetails.Any(static warning => warning.IsPartialData),
            IsUnsupported = warningDetails.Any(static warning => string.Equals(warning.Code, OfflineAddressBookCmdletUnavailableCode, StringComparison.Ordinal))
        };
    }

    public async Task UpsertOfflineAddressBookAsync(UpsertOfflineAddressBookRequest request, CancellationToken cancellationToken = default)
    {
        var script = OrganizationDirectoryCommandBuilder.BuildUpsertOfflineAddressBookScript(request);
        await RunScriptAsync(script, cancellationToken);
    }

    public async Task RemoveOfflineAddressBookAsync(RemoveOfflineAddressBookRequest request, CancellationToken cancellationToken = default)
    {
        var identity = EscapePs(request.Identity);
        var script = $@"
Remove-OfflineAddressBook -Identity '{identity}' -Confirm:$false -ErrorAction Stop
Write-Output 'OK'";

        await RunScriptAsync(script, cancellationToken);
    }
}

