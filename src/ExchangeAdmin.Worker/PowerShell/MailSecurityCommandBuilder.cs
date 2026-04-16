using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

internal static class MailSecurityCommandBuilder
{
    public static string BuildUpdateDkimSigningConfigScript(UpdateDkimSigningConfigRequest request)
    {
        var identity = EscapePs(request.Identity);
        var enabled = ToPsBoolLiteral(request.Enabled);

        return $@"
Set-DkimSigningConfig -Identity '{identity}' -Enabled {enabled} -ErrorAction Stop
Write-Output 'OK'";
    }

    public static string BuildUpdateHostedContentFilterPolicyScript(UpdateHostedContentFilterPolicyRequest request)
    {
        var identity = EscapePs(request.Identity);
        var ruleIdentity = EscapePs(request.RuleIdentity);
        var enabledLiteral = request.Enabled.HasValue ? ToPsBoolLiteral(request.Enabled.Value) : "$null";
        var spamAction = EscapePs(request.SpamAction);
        var highConfidenceSpamAction = EscapePs(request.HighConfidenceSpamAction);
        var phishSpamAction = EscapePs(request.PhishSpamAction);
        var bulkThreshold = request.BulkThreshold.HasValue ? request.BulkThreshold.Value.ToString() : "$null";

        return $@"
$params = @{{}}
if ({bulkThreshold} -ne $null) {{ $params['BulkThreshold'] = {bulkThreshold} }}
if (![string]::IsNullOrWhiteSpace('{spamAction}')) {{ $params['SpamAction'] = '{spamAction}' }}
if (![string]::IsNullOrWhiteSpace('{highConfidenceSpamAction}')) {{ $params['HighConfidenceSpamAction'] = '{highConfidenceSpamAction}' }}
if (![string]::IsNullOrWhiteSpace('{phishSpamAction}')) {{ $params['PhishSpamAction'] = '{phishSpamAction}' }}
if ($params.Count -gt 0) {{
    Set-HostedContentFilterPolicy -Identity '{identity}' @params -ErrorAction Stop
}}

if (!([string]::IsNullOrWhiteSpace('{ruleIdentity}')) -and {enabledLiteral} -ne $null) {{
    if ({enabledLiteral}) {{
        Enable-HostedContentFilterRule -Identity '{ruleIdentity}' -Confirm:$false -ErrorAction Stop
    }}
    else {{
        Disable-HostedContentFilterRule -Identity '{ruleIdentity}' -Confirm:$false -ErrorAction Stop
    }}
}}

Write-Output 'OK'";
    }

    public static string BuildUpdateAntiPhishPolicyScript(UpdateAntiPhishPolicyRequest request)
    {
        var identity = EscapePs(request.Identity);
        var ruleIdentity = EscapePs(request.RuleIdentity);
        var enabledLiteral = request.Enabled.HasValue ? ToPsBoolLiteral(request.Enabled.Value) : "$null";
        var mailboxAction = EscapePs(request.MailboxIntelligenceProtectionAction);
        var targetedAction = EscapePs(request.TargetedUserProtectionAction);
        var authenticationFailAction = EscapePs(request.AuthenticationFailAction);
        var dmarcRejectAction = EscapePs(request.DmarcRejectAction);
        var dmarcQuarantineAction = EscapePs(request.DmarcQuarantineAction);
        var phishThresholdLevel = request.PhishThresholdLevel.HasValue ? request.PhishThresholdLevel.Value.ToString() : "$null";

        return $@"
$params = @{{}}
if ({phishThresholdLevel} -ne $null) {{ $params['PhishThresholdLevel'] = {phishThresholdLevel} }}
if ({ToPsNullableBoolLiteral(request.EnableSpoofIntelligence)} -ne $null) {{ $params['EnableSpoofIntelligence'] = {ToPsNullableBoolLiteral(request.EnableSpoofIntelligence)} }}
if ({ToPsNullableBoolLiteral(request.EnableMailboxIntelligence)} -ne $null) {{ $params['EnableMailboxIntelligence'] = {ToPsNullableBoolLiteral(request.EnableMailboxIntelligence)} }}
if ({ToPsNullableBoolLiteral(request.EnableTargetedUserProtection)} -ne $null) {{ $params['EnableTargetedUserProtection'] = {ToPsNullableBoolLiteral(request.EnableTargetedUserProtection)} }}
if ({ToPsNullableBoolLiteral(request.HonorDmarcPolicy)} -ne $null) {{ $params['HonorDmarcPolicy'] = {ToPsNullableBoolLiteral(request.HonorDmarcPolicy)} }}
if (![string]::IsNullOrWhiteSpace('{mailboxAction}')) {{ $params['MailboxIntelligenceProtectionAction'] = '{mailboxAction}' }}
if (![string]::IsNullOrWhiteSpace('{targetedAction}')) {{ $params['TargetedUserProtectionAction'] = '{targetedAction}' }}
if (![string]::IsNullOrWhiteSpace('{authenticationFailAction}')) {{ $params['AuthenticationFailAction'] = '{authenticationFailAction}' }}
if (![string]::IsNullOrWhiteSpace('{dmarcRejectAction}')) {{ $params['DmarcRejectAction'] = '{dmarcRejectAction}' }}
if (![string]::IsNullOrWhiteSpace('{dmarcQuarantineAction}')) {{ $params['DmarcQuarantineAction'] = '{dmarcQuarantineAction}' }}
if ($params.Count -gt 0) {{
    Set-AntiPhishPolicy -Identity '{identity}' @params -ErrorAction Stop
}}

if (!([string]::IsNullOrWhiteSpace('{ruleIdentity}')) -and {enabledLiteral} -ne $null) {{
    if ({enabledLiteral}) {{
        Enable-AntiPhishRule -Identity '{ruleIdentity}' -Confirm:$false -ErrorAction Stop
    }}
    else {{
        Disable-AntiPhishRule -Identity '{ruleIdentity}' -Confirm:$false -ErrorAction Stop
    }}
}}

Write-Output 'OK'";
    }

    public static string BuildUpdateMalwareFilterPolicyScript(UpdateMalwareFilterPolicyRequest request)
    {
        var identity = EscapePs(request.Identity);
        var ruleIdentity = EscapePs(request.RuleIdentity);
        var enabledLiteral = request.Enabled.HasValue ? ToPsBoolLiteral(request.Enabled.Value) : "$null";
        var fileTypeAction = EscapePs(request.FileTypeAction);

        return $@"
$params = @{{}}
if ({ToPsNullableBoolLiteral(request.EnableFileFilter)} -ne $null) {{ $params['EnableFileFilter'] = {ToPsNullableBoolLiteral(request.EnableFileFilter)} }}
if ({ToPsNullableBoolLiteral(request.ZapEnabled)} -ne $null) {{ $params['ZapEnabled'] = {ToPsNullableBoolLiteral(request.ZapEnabled)} }}
if (![string]::IsNullOrWhiteSpace('{fileTypeAction}')) {{ $params['FileTypeAction'] = '{fileTypeAction}' }}
if ($params.Count -gt 0) {{
    Set-MalwareFilterPolicy -Identity '{identity}' @params -ErrorAction Stop
}}

if (!([string]::IsNullOrWhiteSpace('{ruleIdentity}')) -and {enabledLiteral} -ne $null) {{
    if ({enabledLiteral}) {{
        Enable-MalwareFilterRule -Identity '{ruleIdentity}' -Confirm:$false -ErrorAction Stop
    }}
    else {{
        Disable-MalwareFilterRule -Identity '{ruleIdentity}' -Confirm:$false -ErrorAction Stop
    }}
}}

Write-Output 'OK'";
    }

    public static string BuildUpdateQuarantinePolicyScript(UpdateQuarantinePolicyRequest request)
    {
        var identity = EscapePs(request.Identity);
        var permissions = EscapePs(request.EndUserQuarantinePermissionsValue);
        var frequency = request.EndUserSpamNotificationFrequency.HasValue ? request.EndUserSpamNotificationFrequency.Value.ToString() : "$null";
        var retention = request.QuarantineRetentionDays.HasValue ? request.QuarantineRetentionDays.Value.ToString() : "$null";

        return $@"
$params = @{{}}
if (![string]::IsNullOrWhiteSpace('{permissions}')) {{ $params['EndUserQuarantinePermissionsValue'] = '{permissions}' }}
if ({ToPsNullableBoolLiteral(request.EsnEnabled)} -ne $null) {{ $params['ESNEnabled'] = {ToPsNullableBoolLiteral(request.EsnEnabled)} }}
if ({frequency} -ne $null) {{ $params['EndUserSpamNotificationFrequency'] = {frequency} }}
if ({retention} -ne $null) {{ $params['QuarantineRetentionDays'] = {retention} }}
if ({ToPsNullableBoolLiteral(request.OrganizationBrandingEnabled)} -ne $null) {{ $params['OrganizationBrandingEnabled'] = {ToPsNullableBoolLiteral(request.OrganizationBrandingEnabled)} }}
if ($params.Count -eq 0) {{
    Write-Output 'OK'
    return
}}

Set-QuarantinePolicy -Identity '{identity}' @params -ErrorAction Stop
Write-Output 'OK'";
    }

    public static string BuildUpdateHostedOutboundSpamFilterPolicyScript(UpdateHostedOutboundSpamFilterPolicyRequest request)
    {
        var identity = EscapePs(request.Identity);
        var actionWhenThresholdReached = EscapePs(request.ActionWhenThresholdReached);
        var autoForwardingMode = EscapePs(request.AutoForwardingMode);
        var recipientLimitExternalPerHour = request.RecipientLimitExternalPerHour.HasValue ? request.RecipientLimitExternalPerHour.Value.ToString() : "$null";
        var recipientLimitInternalPerHour = request.RecipientLimitInternalPerHour.HasValue ? request.RecipientLimitInternalPerHour.Value.ToString() : "$null";
        var recipientLimitPerDay = request.RecipientLimitPerDay.HasValue ? request.RecipientLimitPerDay.Value.ToString() : "$null";

        return $@"
$params = @{{}}
if ({recipientLimitExternalPerHour} -ne $null) {{ $params['RecipientLimitExternalPerHour'] = {recipientLimitExternalPerHour} }}
if ({recipientLimitInternalPerHour} -ne $null) {{ $params['RecipientLimitInternalPerHour'] = {recipientLimitInternalPerHour} }}
if ({recipientLimitPerDay} -ne $null) {{ $params['RecipientLimitPerDay'] = {recipientLimitPerDay} }}
if (![string]::IsNullOrWhiteSpace('{actionWhenThresholdReached}')) {{ $params['ActionWhenThresholdReached'] = '{actionWhenThresholdReached}' }}
if (![string]::IsNullOrWhiteSpace('{autoForwardingMode}')) {{ $params['AutoForwardingMode'] = '{autoForwardingMode}' }}
if ($params.Count -eq 0) {{
    Write-Output 'OK'
    return
}}

Set-HostedOutboundSpamFilterPolicy -Identity '{identity}' @params -ErrorAction Stop
Write-Output 'OK'";
    }

    private static string EscapePs(string? value)
        => (value ?? string.Empty).Replace("'", "''");

    private static string ToPsBoolLiteral(bool value)
        => value ? "$true" : "$false";

    private static string ToPsNullableBoolLiteral(bool? value)
        => value.HasValue ? ToPsBoolLiteral(value.Value) : "$null";
}
