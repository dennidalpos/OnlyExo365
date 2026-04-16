using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

internal sealed class ExoRemoteDomainCommands : ExoCommandModuleBase
{
    public ExoRemoteDomainCommands(PowerShellEngine engine)
        : base(engine)
    {
    }

    public async Task<GetRemoteDomainsResponse> GetRemoteDomainsAsync(CancellationToken cancellationToken = default)
    {
        var script = @"
Get-RemoteDomain -ErrorAction Stop |
    Sort-Object Name |
    ForEach-Object {
        $domainName = if ($_.DomainName) { $_.DomainName.ToString() } else { '' }
        [PSCustomObject]@{
            Identity = $_.Identity.ToString()
            Name = $_.Name
            DomainName = $domainName
            AllowedOOFType = if ($_.AllowedOOFType) { $_.AllowedOOFType.ToString() } else { 'External' }
            AutoReplyEnabled = [bool]$_.AutoReplyEnabled
            AutoForwardEnabled = [bool]$_.AutoForwardEnabled
            DeliveryReportEnabled = [bool]$_.DeliveryReportEnabled
            NDREnabled = [bool]$_.NDREnabled
            MeetingForwardNotificationEnabled = [bool]$_.MeetingForwardNotificationEnabled
            TNEFEnabled = [bool]$_.TNEFEnabled
            TrustedMailOutboundEnabled = [bool]$_.TrustedMailOutboundEnabled
            IsDefault = ($domainName -eq '*' -or $_.Identity.ToString() -eq 'Default')
        }
    }";

        var results = await RunScriptAsync(script, cancellationToken);
        var domains = new List<RemoteDomainDto>();

        foreach (var obj in results)
        {
            domains.Add(new RemoteDomainDto
            {
                Identity = GetString(obj, "Identity"),
                Name = GetString(obj, "Name"),
                DomainName = GetString(obj, "DomainName"),
                AllowedOOFType = GetString(obj, "AllowedOOFType"),
                AutoReplyEnabled = GetBool(obj, "AutoReplyEnabled"),
                AutoForwardEnabled = GetBool(obj, "AutoForwardEnabled"),
                DeliveryReportEnabled = GetBool(obj, "DeliveryReportEnabled"),
                NDREnabled = GetBool(obj, "NDREnabled"),
                MeetingForwardNotificationEnabled = GetBool(obj, "MeetingForwardNotificationEnabled"),
                TNEFEnabled = GetBool(obj, "TNEFEnabled"),
                TrustedMailOutboundEnabled = GetBool(obj, "TrustedMailOutboundEnabled"),
                IsDefault = GetBool(obj, "IsDefault")
            });
        }

        return new GetRemoteDomainsResponse
        {
            Domains = domains
        };
    }

    public async Task UpsertRemoteDomainAsync(UpsertRemoteDomainRequest request, CancellationToken cancellationToken = default)
    {
        request.AllowedOOFType = NormalizeAllowedOofType(request.AllowedOOFType);
        var script = RemoteDomainCommandBuilder.BuildUpsertRemoteDomainScript(request);
        await RunScriptAsync(script, cancellationToken);
    }

    public async Task RemoveRemoteDomainAsync(RemoveRemoteDomainRequest request, CancellationToken cancellationToken = default)
    {
        var identity = EscapePs(request.Identity);
        var script = $@"
Remove-RemoteDomain -Identity '{identity}' -Confirm:$false -ErrorAction Stop
Write-Output 'OK'";

        await RunScriptAsync(script, cancellationToken);
    }

    private static string NormalizeAllowedOofType(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "External";
        }

        return value.Trim() switch
        {
            "External" => "External",
            "ExternalLegacy" => "ExternalLegacy",
            "InternalLegacy" => "InternalLegacy",
            "None" => "None",
            _ => "External"
        };
    }
}
