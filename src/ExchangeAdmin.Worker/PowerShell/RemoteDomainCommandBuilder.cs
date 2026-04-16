using System.Text;
using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

public static class RemoteDomainCommandBuilder
{
    public static string BuildUpsertRemoteDomainScript(UpsertRemoteDomainRequest request)
    {
        var escapedIdentity = EscapePs(request.Identity);
        var escapedName = EscapePs(request.Name);
        var escapedDomainName = EscapePs(request.DomainName);
        var escapedAllowedOofType = EscapePs(string.IsNullOrWhiteSpace(request.AllowedOOFType) ? "External" : request.AllowedOOFType);

        var script = new StringBuilder();
        script.AppendLine("$params = @{");
        script.AppendLine($"    AllowedOOFType = '{escapedAllowedOofType}'");
        script.AppendLine($"    AutoReplyEnabled = {ToPsBoolLiteral(request.AutoReplyEnabled)}");
        script.AppendLine($"    AutoForwardEnabled = {ToPsBoolLiteral(request.AutoForwardEnabled)}");
        script.AppendLine($"    DeliveryReportEnabled = {ToPsBoolLiteral(request.DeliveryReportEnabled)}");
        script.AppendLine($"    NDREnabled = {ToPsBoolLiteral(request.NDREnabled)}");
        script.AppendLine($"    MeetingForwardNotificationEnabled = {ToPsBoolLiteral(request.MeetingForwardNotificationEnabled)}");
        script.AppendLine($"    TNEFEnabled = {ToPsBoolLiteral(request.TNEFEnabled)}");
        script.AppendLine($"    TrustedMailOutboundEnabled = {ToPsBoolLiteral(request.TrustedMailOutboundEnabled)}");
        script.AppendLine("}");
        script.AppendLine();
        script.AppendLine($"if ('{escapedIdentity}' -ne '') {{");
        script.AppendLine($"    if ('{escapedName}' -ne '') {{ $params['Name'] = '{escapedName}' }}");
        script.AppendLine($"    Set-RemoteDomain -Identity '{escapedIdentity}' @params -ErrorAction Stop");
        script.AppendLine("} else {");
        script.AppendLine($"    New-RemoteDomain -Name '{escapedName}' -DomainName '{escapedDomainName}' -ErrorAction Stop | Out-Null");
        script.AppendLine($"    Set-RemoteDomain -Identity '{escapedName}' @params -ErrorAction Stop");
        script.AppendLine("}");

        return script.ToString();
    }

    private static string EscapePs(string? value)
        => (value ?? string.Empty).Replace("'", "''");

    private static string ToPsBoolLiteral(bool value)
        => value ? "$true" : "$false";
}
