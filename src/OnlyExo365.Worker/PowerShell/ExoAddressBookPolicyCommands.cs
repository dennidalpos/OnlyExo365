using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

internal sealed class ExoAddressBookPolicyCommands : ExoCommandModuleBase
{
    public ExoAddressBookPolicyCommands(PowerShellEngine engine)
        : base(engine)
    {
    }

    public async Task<GetAddressBookPoliciesResponse> GetAddressBookPoliciesAsync(CancellationToken cancellationToken = default)
    {
        var script = @"
Get-AddressBookPolicy -ErrorAction Stop |
    Sort-Object Name |
    ForEach-Object {
        [PSCustomObject]@{
            Identity = $_.Identity.ToString()
            Name = $_.Name
            AddressLists = @($_.AddressLists | ForEach-Object { $_.ToString() })
            GlobalAddressList = if ($_.GlobalAddressList) { $_.GlobalAddressList.ToString() } else { '' }
            OfflineAddressBook = if ($_.OfflineAddressBook) { $_.OfflineAddressBook.ToString() } else { '' }
            RoomList = if ($_.RoomList) { $_.RoomList.ToString() } else { '' }
        }
    }";

        var results = await RunScriptAsync(script, cancellationToken);
        var items = new List<AddressBookPolicyDto>();

        foreach (var obj in results)
        {
            items.Add(new AddressBookPolicyDto
            {
                Identity = GetString(obj, "Identity"),
                Name = GetString(obj, "Name"),
                AddressLists = ConvertToStringList(obj.Properties["AddressLists"]?.Value),
                GlobalAddressList = GetString(obj, "GlobalAddressList"),
                OfflineAddressBook = GetString(obj, "OfflineAddressBook"),
                RoomList = GetString(obj, "RoomList")
            });
        }

        return new GetAddressBookPoliciesResponse
        {
            Policies = items
        };
    }

    public async Task UpsertAddressBookPolicyAsync(UpsertAddressBookPolicyRequest request, CancellationToken cancellationToken = default)
    {
        var script = OrganizationDirectoryCommandBuilder.BuildUpsertAddressBookPolicyScript(request);
        await RunScriptAsync(script, cancellationToken);
    }

    public async Task RemoveAddressBookPolicyAsync(RemoveAddressBookPolicyRequest request, CancellationToken cancellationToken = default)
    {
        var identity = EscapePs(request.Identity);
        var script = $@"
Remove-AddressBookPolicy -Identity '{identity}' -Confirm:$false -ErrorAction Stop
Write-Output 'OK'";

        await RunScriptAsync(script, cancellationToken);
    }
}

