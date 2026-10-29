using System.Text;
using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

public static class OrganizationRelationshipCommandBuilder
{
    public static string BuildUpsertOrganizationRelationshipScript(UpsertOrganizationRelationshipRequest request)
    {
        var escapedIdentity = EscapePs(request.Identity);
        var escapedName = EscapePs(request.Name);
        var escapedTargetApplicationUri = EscapePs(request.TargetApplicationUri);
        var escapedTargetAutodiscoverEpr = EscapePs(request.TargetAutodiscoverEpr);
        var escapedFreeBusyAccessLevel = EscapePs(string.IsNullOrWhiteSpace(request.FreeBusyAccessLevel) ? "AvailabilityOnly" : request.FreeBusyAccessLevel);
        var escapedMailTipsAccessLevel = EscapePs(string.IsNullOrWhiteSpace(request.MailTipsAccessLevel) ? "All" : request.MailTipsAccessLevel);

        var script = new StringBuilder();
        script.AppendLine($"$domainNames = {ToPsArrayLiteral(request.DomainNames)}");
        script.AppendLine("$params = @{");
        script.AppendLine($"    DomainNames = $domainNames");
        script.AppendLine($"    Enabled = {ToPsBoolLiteral(request.Enabled)}");
        script.AppendLine($"    FreeBusyAccessEnabled = {ToPsBoolLiteral(request.FreeBusyAccessEnabled)}");
        script.AppendLine($"    FreeBusyAccessLevel = '{escapedFreeBusyAccessLevel}'");
        script.AppendLine($"    MailTipsAccessEnabled = {ToPsBoolLiteral(request.MailTipsAccessEnabled)}");
        script.AppendLine($"    MailTipsAccessLevel = '{escapedMailTipsAccessLevel}'");
        AppendNullableBool(script, "ArchiveAccessEnabled", request.ArchiveAccessEnabled);
        AppendNullableBool(script, "DeliveryReportEnabled", request.DeliveryReportEnabled);
        AppendNullableBool(script, "MailboxMoveEnabled", request.MailboxMoveEnabled);
        AppendNullableBool(script, "PhotosEnabled", request.PhotosEnabled);
        script.AppendLine("}");
        script.AppendLine();
        script.AppendLine($"if ('{escapedTargetApplicationUri}' -ne '') {{ $params['TargetApplicationUri'] = '{escapedTargetApplicationUri}' }}");
        script.AppendLine($"if ('{escapedTargetAutodiscoverEpr}' -ne '') {{ $params['TargetAutodiscoverEpr'] = '{escapedTargetAutodiscoverEpr}' }}");
        script.AppendLine();
        script.AppendLine($"if ('{escapedIdentity}' -ne '') {{");
        script.AppendLine($"    Set-OrganizationRelationship -Identity '{escapedIdentity}' @params -ErrorAction Stop");
        script.AppendLine("} else {");
        script.AppendLine($"    New-OrganizationRelationship -Name '{escapedName}' @params -ErrorAction Stop | Out-Null");
        script.AppendLine("}");

        return script.ToString();
    }

    private static void AppendNullableBool(StringBuilder script, string name, bool? value)
    {
        if (value.HasValue)
        {
            script.AppendLine($"    {name} = {ToPsBoolLiteral(value.Value)}");
        }
    }

    private static string EscapePs(string? value)
        => (value ?? string.Empty).Replace("'", "''");

    private static string ToPsBoolLiteral(bool value)
        => value ? "$true" : "$false";

    private static string ToPsArrayLiteral(IEnumerable<string>? values)
    {
        if (values == null)
        {
            return "@()";
        }

        var normalized = values
            .Where(static value => !string.IsNullOrWhiteSpace(value))
            .Select(static value => $"'{EscapePs(value.Trim())}'")
            .ToArray();

        return normalized.Length == 0
            ? "@()"
            : "@(" + string.Join(", ", normalized) + ")";
    }
}

