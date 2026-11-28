using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

internal sealed class ExoAddressListCommands : ExoCommandModuleBase
{
    private const string AddressListCmdletUnavailableCode = "AddressListCmdletUnavailable";

    public ExoAddressListCommands(PowerShellEngine engine)
        : base(engine)
    {
    }

    public async Task<GetAddressListsResponse> GetAddressListsAsync(CancellationToken cancellationToken = default)
    {
        var script = BuildGetAddressListsScript();
        var result = await Engine.ExecuteAsync(script, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException(result.ErrorMessage ?? "Address list query failed");
        }

        return MapAddressListsResponse(result.Output, result.Warning);
    }

    internal static string BuildGetAddressListsScript()
    {
        return @"
if (-not (Get-Command -Name 'Get-AddressList' -ErrorAction SilentlyContinue)) {
    $warningPayload = @{
        Code = '__CODE__'
        Scope = 'MailFlow.AddressLists'
        Message = 'Address Lists are not supported in the current Exchange session: Get-AddressList is not available.'
        IsPartialData = $true
    } | ConvertTo-Json -Compress -Depth 3
    Write-Warning ('__WARN__' + $warningPayload)
    return
}

Get-AddressList -ErrorAction Stop |
    Sort-Object DisplayName, Name |
    ForEach-Object {
        [PSCustomObject]@{
            Identity = $_.Identity.ToString()
            Name = $_.Name
            DisplayName = if ($_.DisplayName) { $_.DisplayName } else { $_.Name }
            RecipientFilter = if ($_.RecipientFilter) { $_.RecipientFilter.ToString() } else { '' }
            RecipientContainer = if ($_.RecipientContainer) { $_.RecipientContainer.ToString() } else { $null }
            IncludedRecipients = @($_.IncludedRecipients | ForEach-Object { $_.ToString() })
            ConditionalCompany = @($_.ConditionalCompany | ForEach-Object { $_.ToString() })
            ConditionalDepartment = @($_.ConditionalDepartment | ForEach-Object { $_.ToString() })
            ConditionalStateOrProvince = @($_.ConditionalStateOrProvince | ForEach-Object { $_.ToString() })
            ConditionalCustomAttribute1 = @($_.ConditionalCustomAttribute1 | ForEach-Object { $_.ToString() })
        }
    }"
            .Replace("__CODE__", AddressListCmdletUnavailableCode, StringComparison.Ordinal)
            .Replace("__WARN__", StructuredWarningPrefix, StringComparison.Ordinal);
    }

    internal static GetAddressListsResponse MapAddressListsResponse(
        IEnumerable<System.Management.Automation.PSObject> results,
        IEnumerable<string>? warnings)
    {
        var warningDetails = ParseStructuredWarnings(warnings);
        var items = new List<AddressListDto>();

        foreach (var obj in results)
        {
            items.Add(new AddressListDto
            {
                Identity = GetString(obj, "Identity"),
                Name = GetString(obj, "Name"),
                DisplayName = GetString(obj, "DisplayName"),
                RecipientFilter = GetString(obj, "RecipientFilter"),
                RecipientContainer = GetNullableString(obj, "RecipientContainer"),
                IncludedRecipients = ConvertToStringList(obj.Properties["IncludedRecipients"]?.Value),
                ConditionalCompany = ConvertToStringList(obj.Properties["ConditionalCompany"]?.Value),
                ConditionalDepartment = ConvertToStringList(obj.Properties["ConditionalDepartment"]?.Value),
                ConditionalStateOrProvince = ConvertToStringList(obj.Properties["ConditionalStateOrProvince"]?.Value),
                ConditionalCustomAttribute1 = ConvertToStringList(obj.Properties["ConditionalCustomAttribute1"]?.Value)
            });
        }

        return new GetAddressListsResponse
        {
            AddressLists = items,
            Warnings = ExtractWarningMessages(warningDetails),
            WarningDetails = warningDetails,
            HasPartialData = warningDetails.Any(static warning => warning.IsPartialData),
            IsUnsupported = warningDetails.Any(static warning => string.Equals(warning.Code, AddressListCmdletUnavailableCode, StringComparison.Ordinal))
        };
    }

    public async Task UpsertAddressListAsync(UpsertAddressListRequest request, CancellationToken cancellationToken = default)
    {
        var script = OrganizationDirectoryCommandBuilder.BuildUpsertAddressListScript(request);
        await RunScriptAsync(script, cancellationToken);
    }

    public async Task RemoveAddressListAsync(RemoveAddressListRequest request, CancellationToken cancellationToken = default)
    {
        var identity = EscapePs(request.Identity);
        var script = $@"
Remove-AddressList -Identity '{identity}' -Confirm:$false -ErrorAction Stop
Write-Output 'OK'";

        await RunScriptAsync(script, cancellationToken);
    }
}

