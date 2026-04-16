using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

internal sealed class ExoAcceptedDomainCommands : ExoCommandModuleBase
{
    public ExoAcceptedDomainCommands(PowerShellEngine engine)
        : base(engine)
    {
    }

    public async Task<GetAcceptedDomainsResponse> GetAcceptedDomainsAsync(CancellationToken cancellationToken = default)
    {
        await EnsureExchangeCmdletAvailableAsync("Get-AcceptedDomain", cancellationToken);

        var script = @"
Get-AcceptedDomain -ErrorAction Stop | ForEach-Object {
    [PSCustomObject]@{
        Identity = $_.Identity.ToString()
        Name = $_.Name
        DomainName = $_.DomainName.ToString()
        DomainType = $_.DomainType.ToString()
        Default = [bool]$_.Default
    }
}";

        var results = await RunScriptAsync(script, cancellationToken);
        var domains = new List<AcceptedDomainDto>();

        foreach (var obj in results)
        {
            domains.Add(new AcceptedDomainDto
            {
                Identity = GetString(obj, "Identity"),
                Name = GetString(obj, "Name"),
                DomainName = GetString(obj, "DomainName"),
                DomainType = GetString(obj, "DomainType"),
                Default = GetBool(obj, "Default")
            });
        }

        return new GetAcceptedDomainsResponse
        {
            Domains = domains.OrderBy(static domain => domain.DomainName).ToList()
        };
    }

    public async Task UpsertAcceptedDomainAsync(UpsertAcceptedDomainRequest request, CancellationToken cancellationToken = default)
    {
        await EnsureExchangeCmdletAvailableAsync(
            string.IsNullOrWhiteSpace(request.Identity) ? "New-AcceptedDomain" : "Set-AcceptedDomain",
            cancellationToken);

        var identity = EscapePs(request.Identity);
        var name = EscapePs(request.Name);
        var domainName = EscapePs(request.DomainName);
        var domainType = EscapePs(request.DomainType);

        var script = $@"
if ('{identity}' -ne '') {{
    Set-AcceptedDomain -Identity '{identity}' -DomainType '{domainType}' -ErrorAction Stop
}} else {{
    New-AcceptedDomain -Name '{name}' -DomainName '{domainName}' -DomainType '{domainType}' -ErrorAction Stop
}}
if {(request.MakeDefault ? "$true" : "$false")} {{
    Set-AcceptedDomain -Identity '{domainName}' -MakeDefault -ErrorAction Stop
}}";

        await RunScriptAsync(script, cancellationToken);
    }

    public async Task RemoveAcceptedDomainAsync(RemoveAcceptedDomainRequest request, CancellationToken cancellationToken = default)
    {
        await EnsureExchangeCmdletAvailableAsync("Remove-AcceptedDomain", cancellationToken);

        var identity = EscapePs(request.Identity);
        var script = $@"
Remove-AcceptedDomain -Identity '{identity}' -Confirm:$false -ErrorAction Stop
Write-Output 'OK'";

        await RunScriptAsync(script, cancellationToken);
    }
}
