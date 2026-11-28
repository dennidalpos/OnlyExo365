using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

internal sealed class ExoSharingPolicyCommands : ExoCommandModuleBase
{
    public ExoSharingPolicyCommands(PowerShellEngine engine)
        : base(engine)
    {
    }

    public async Task<GetSharingPoliciesResponse> GetSharingPoliciesAsync(CancellationToken cancellationToken = default)
    {
        var script = @"
Get-SharingPolicy -ErrorAction Stop |
    Sort-Object Name |
    ForEach-Object {
        [PSCustomObject]@{
            Identity = $_.Identity.ToString()
            Name = $_.Name
            Domains = @($_.Domains | ForEach-Object { $_.ToString() })
            Enabled = [bool]$_.Enabled
            IsDefault = if ($null -ne $_.IsDefault) { [bool]$_.IsDefault } elseif ($null -ne $_.Default) { [bool]$_.Default } else { $false }
        }
    }";

        var results = await RunScriptAsync(script, cancellationToken);
        var items = new List<SharingPolicyDto>();

        foreach (var obj in results)
        {
            items.Add(new SharingPolicyDto
            {
                Identity = GetString(obj, "Identity"),
                Name = GetString(obj, "Name"),
                Domains = ConvertToStringList(obj.Properties["Domains"]?.Value),
                Enabled = GetBool(obj, "Enabled"),
                IsDefault = GetBool(obj, "IsDefault")
            });
        }

        return new GetSharingPoliciesResponse
        {
            Policies = items
        };
    }

    public async Task UpsertSharingPolicyAsync(UpsertSharingPolicyRequest request, CancellationToken cancellationToken = default)
    {
        var script = OrganizationDirectoryCommandBuilder.BuildUpsertSharingPolicyScript(request);
        await RunScriptAsync(script, cancellationToken);
    }

    public async Task RemoveSharingPolicyAsync(RemoveSharingPolicyRequest request, CancellationToken cancellationToken = default)
    {
        var identity = EscapePs(request.Identity);
        var script = $@"
Remove-SharingPolicy -Identity '{identity}' -Confirm:$false -ErrorAction Stop
Write-Output 'OK'";

        await RunScriptAsync(script, cancellationToken);
    }
}

