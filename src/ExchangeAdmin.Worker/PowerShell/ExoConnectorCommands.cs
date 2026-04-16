using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

internal sealed class ExoConnectorCommands : ExoCommandModuleBase
{
    public ExoConnectorCommands(PowerShellEngine engine)
        : base(engine)
    {
    }

    public async Task<GetConnectorsResponse> GetConnectorsAsync(CancellationToken cancellationToken = default)
    {
        var script = @"
$inbound = Get-InboundConnector -ErrorAction SilentlyContinue | ForEach-Object {
    [PSCustomObject]@{
        Identity = $_.Identity.ToString()
        Name = $_.Name
        Type = 'Inbound'
        Enabled = [bool]$_.Enabled
        Comment = if ($_.Comment) { $_.Comment } else { '' }
        SenderDomains = @($_.SenderDomains)
        RecipientDomains = @($_.RecipientDomains)
    }
}
$outbound = Get-OutboundConnector -ErrorAction SilentlyContinue | ForEach-Object {
    [PSCustomObject]@{
        Identity = $_.Identity.ToString()
        Name = $_.Name
        Type = 'Outbound'
        Enabled = [bool]$_.Enabled
        Comment = if ($_.Comment) { $_.Comment } else { '' }
        SenderDomains = @($_.SenderDomains)
        RecipientDomains = @($_.RecipientDomains)
    }
}
@($inbound + $outbound)";

        var results = await RunScriptAllowErrorsAsync(script, cancellationToken: cancellationToken);
        var connectors = new List<ConnectorDto>();

        foreach (var obj in results)
        {
            connectors.Add(new ConnectorDto
            {
                Identity = GetString(obj, "Identity"),
                Name = GetString(obj, "Name"),
                DisplayLabel = NormalizeConnectorDisplayLabel(GetString(obj, "Name"), GetString(obj, "Type")),
                Type = GetString(obj, "Type"),
                Enabled = GetBool(obj, "Enabled"),
                Comment = GetString(obj, "Comment"),
                SenderDomains = ConvertToStringList(obj.Properties["SenderDomains"]?.Value),
                RecipientDomains = ConvertToStringList(obj.Properties["RecipientDomains"]?.Value)
            });
        }

        return new GetConnectorsResponse
        {
            Connectors = connectors.OrderBy(static connector => connector.Type).ThenBy(static connector => connector.Name).ToList()
        };
    }

    public async Task UpsertConnectorAsync(UpsertConnectorRequest request, CancellationToken cancellationToken = default)
    {
        var identity = EscapePs(request.Identity);
        var name = EscapePs(request.Name);
        var senderDomains = ToPsArrayLiteral(request.SenderDomains);
        var recipientDomains = ToPsArrayLiteral(request.RecipientDomains);
        var isInbound = string.Equals(request.Type, "Inbound", StringComparison.OrdinalIgnoreCase);

        var script = $@"
$senderDomains = {senderDomains}
$recipientDomains = {recipientDomains}
$enabled = {ToPsBoolLiteral(request.Enabled)}
$comment = {FormatNullableString(request.Comment)}

if {(isInbound ? "$true" : "$false")} {{
    $params = @{{
        Name = '{name}'
        Enabled = $enabled
        ConnectorType = 'OnPremises'
    }}

    if ($comment -ne $null) {{ $params['Comment'] = $comment }}
    if ($senderDomains.Count -gt 0) {{ $params['SenderDomains'] = $senderDomains }}
    if ($recipientDomains.Count -gt 0) {{
        Write-Warning 'RecipientDomains not applicable to Inbound connector and will be ignored.'
    }}

    if ('{identity}' -ne '') {{
        Set-InboundConnector -Identity '{identity}' @params -ErrorAction Stop
    }} else {{
        New-InboundConnector @params -ErrorAction Stop
    }}
}} else {{
    $params = @{{
        Name = '{name}'
        Enabled = $enabled
        ConnectorType = 'OnPremises'
    }}

    if ($comment -ne $null) {{ $params['Comment'] = $comment }}

    if ($senderDomains.Count -gt 0) {{
        Write-Warning 'SenderDomains not applicable to Outbound connector and will be ignored.'
    }}

    if ('{identity}' -ne '') {{
        Set-OutboundConnector -Identity '{identity}' @params -RecipientDomains $recipientDomains -ErrorAction Stop
    }} else {{
        New-OutboundConnector @params -RecipientDomains $recipientDomains -ErrorAction Stop
    }}
}}";

        await RunScriptAsync(script, cancellationToken);
    }

    public async Task RemoveConnectorAsync(RemoveConnectorRequest request, CancellationToken cancellationToken = default)
    {
        var identity = EscapePs(request.Identity);
        var isInbound = string.Equals(request.Type, "Inbound", StringComparison.OrdinalIgnoreCase);
        var script = $@"
if {(isInbound ? "$true" : "$false")} {{
    Remove-InboundConnector -Identity '{identity}' -Confirm:$false -ErrorAction Stop
}} else {{
    Remove-OutboundConnector -Identity '{identity}' -Confirm:$false -ErrorAction Stop
}}
Write-Output 'OK'";

        await RunScriptAsync(script, cancellationToken);
    }

    private static string NormalizeConnectorDisplayLabel(string name, string type)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return string.Empty;
        }

        var trimmed = name.Trim();
        foreach (var separator in new[] { " from ", " to " })
        {
            if (!trimmed.Contains(separator, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var parts = trimmed.Split(separator, 2, StringSplitOptions.TrimEntries);
            if (parts.Length == 2 && Guid.TryParse(parts[1], out _))
            {
                return $"{parts[0]} connector ({type})";
            }
        }

        return trimmed;
    }

    private static string FormatNullableString(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "$null";
        }

        return $"'{EscapePs(value)}'";
    }
}
