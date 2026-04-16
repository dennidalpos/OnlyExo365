using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

internal sealed class ExoMailSecurityCommands : ExoCommandModuleBase
{
    public ExoMailSecurityCommands(PowerShellEngine engine)
        : base(engine)
    {
    }

    public async Task<GetMailSecurityBaselineResponse> GetMailSecurityBaselineAsync(
        GetMailSecurityBaselineRequest request,
        CancellationToken cancellationToken = default)
    {
        var response = new GetMailSecurityBaselineResponse();

        response.DkimConfigs = await LoadIfSupportedAsync(
            "Get-DkimSigningConfig",
            LoadDkimSigningConfigsAsync,
            "DKIM",
            response.Warnings,
            cancellationToken);

        response.AntiSpamPolicies = await LoadIfSupportedAsync(
            "Get-HostedContentFilterPolicy",
            LoadHostedContentFilterPoliciesAsync,
            "Anti-spam",
            response.Warnings,
            cancellationToken);

        response.AntiPhishPolicies = await LoadIfSupportedAsync(
            "Get-AntiPhishPolicy",
            LoadAntiPhishPoliciesAsync,
            "Anti-phish",
            response.Warnings,
            cancellationToken);

        response.MalwarePolicies = await LoadIfSupportedAsync(
            "Get-MalwareFilterPolicy",
            LoadMalwareFilterPoliciesAsync,
            "Anti-malware",
            response.Warnings,
            cancellationToken);

        response.QuarantinePolicies = await LoadIfSupportedAsync(
            "Get-QuarantinePolicy",
            LoadQuarantinePoliciesAsync,
            "Quarantine",
            response.Warnings,
            cancellationToken);

        response.OutboundSpamPolicies = await LoadIfSupportedAsync(
            "Get-HostedOutboundSpamFilterPolicy",
            LoadHostedOutboundSpamFilterPoliciesAsync,
            "Outbound spam",
            response.Warnings,
            cancellationToken);

        return response;
    }

    public async Task UpdateDkimSigningConfigAsync(UpdateDkimSigningConfigRequest request, CancellationToken cancellationToken = default)
    {
        await RunScriptAsync(MailSecurityCommandBuilder.BuildUpdateDkimSigningConfigScript(request), cancellationToken);
    }

    public async Task UpdateHostedContentFilterPolicyAsync(UpdateHostedContentFilterPolicyRequest request, CancellationToken cancellationToken = default)
    {
        await RunScriptAsync(MailSecurityCommandBuilder.BuildUpdateHostedContentFilterPolicyScript(request), cancellationToken);
    }

    public async Task UpdateAntiPhishPolicyAsync(UpdateAntiPhishPolicyRequest request, CancellationToken cancellationToken = default)
    {
        await RunScriptAsync(MailSecurityCommandBuilder.BuildUpdateAntiPhishPolicyScript(request), cancellationToken);
    }

    public async Task UpdateMalwareFilterPolicyAsync(UpdateMalwareFilterPolicyRequest request, CancellationToken cancellationToken = default)
    {
        await RunScriptAsync(MailSecurityCommandBuilder.BuildUpdateMalwareFilterPolicyScript(request), cancellationToken);
    }

    public async Task UpdateQuarantinePolicyAsync(UpdateQuarantinePolicyRequest request, CancellationToken cancellationToken = default)
    {
        await RunScriptAsync(MailSecurityCommandBuilder.BuildUpdateQuarantinePolicyScript(request), cancellationToken);
    }

    public async Task UpdateHostedOutboundSpamFilterPolicyAsync(UpdateHostedOutboundSpamFilterPolicyRequest request, CancellationToken cancellationToken = default)
    {
        await RunScriptAsync(MailSecurityCommandBuilder.BuildUpdateHostedOutboundSpamFilterPolicyScript(request), cancellationToken);
    }

    private async Task<List<TItem>> LoadIfSupportedAsync<TItem>(
        string commandName,
        Func<CancellationToken, Task<List<TItem>>> loader,
        string areaName,
        ICollection<string> warnings,
        CancellationToken cancellationToken)
    {
        if (!await IsCommandAvailableAsync(commandName, cancellationToken))
        {
                warnings.Add($"{areaName}: cmdlet {commandName} is not available in the current session.");
            return new List<TItem>();
        }

        try
        {
            return await loader(cancellationToken);
        }
        catch (Exception ex)
        {
            warnings.Add($"{areaName}: {ex.Message}");
            return new List<TItem>();
        }
    }

    private async Task<bool> IsCommandAvailableAsync(string commandName, CancellationToken cancellationToken)
    {
        var escapedName = EscapePs(commandName);
        var script = $@"
$command = Get-Command -Name '{escapedName}' -ErrorAction SilentlyContinue
[PSCustomObject]@{{ Available = ($null -ne $command) }}";

        var results = await RunScriptAllowErrorsAsync(script, cancellationToken: cancellationToken);
        return results.Count > 0 && GetBool(results[0], "Available");
    }

    private async Task<List<DkimSigningConfigDto>> LoadDkimSigningConfigsAsync(CancellationToken cancellationToken)
    {
        var script = @"
Get-DkimSigningConfig -ErrorAction Stop |
    Sort-Object Domain |
    ForEach-Object {
        [PSCustomObject]@{
            Identity = $_.Identity.ToString()
            Domain = if ($_.Domain) { $_.Domain.ToString() } else { $_.Identity.ToString() }
            Enabled = [bool]$_.Enabled
            Selector1CName = if ($_.Selector1CNAME) { $_.Selector1CNAME.ToString() } else { $null }
            Selector2CName = if ($_.Selector2CNAME) { $_.Selector2CNAME.ToString() } else { $null }
            LastChecked = $_.LastChecked
        }
    }";

        var results = await RunScriptAsync(script, cancellationToken);
        return results.Select(obj => new DkimSigningConfigDto
        {
            Identity = GetString(obj, "Identity"),
            Domain = GetString(obj, "Domain"),
            Enabled = GetBool(obj, "Enabled"),
            Selector1CName = GetNullableString(obj, "Selector1CName"),
            Selector2CName = GetNullableString(obj, "Selector2CName"),
            LastChecked = GetNullableDateTime(obj, "LastChecked")
        }).ToList();
    }

    private async Task<List<HostedContentFilterPolicyDto>> LoadHostedContentFilterPoliciesAsync(CancellationToken cancellationToken)
    {
        var script = @"
$rules = @()
if (Get-Command -Name Get-HostedContentFilterRule -ErrorAction SilentlyContinue) {
    $rules = @(Get-HostedContentFilterRule -ErrorAction Stop)
}

Get-HostedContentFilterPolicy -ErrorAction Stop |
    Sort-Object Name |
    ForEach-Object {
        $policy = $_
        $rule = $rules |
            Where-Object { $_.HostedContentFilterPolicy -eq $policy.Name } |
            Sort-Object Priority |
            Select-Object -First 1

        [PSCustomObject]@{
            Identity = $policy.Identity.ToString()
            Name = $policy.Name
            RuleIdentity = if ($rule) { $rule.Identity.ToString() } else { $null }
            RuleName = if ($rule) { $rule.Name } else { $null }
            RuleState = if ($rule) { $rule.State.ToString() } else { $null }
            Priority = if ($rule -and $rule.Priority -ne $null) { [int]$rule.Priority } else { $null }
            BulkThreshold = if ($policy.BulkThreshold -ne $null) { [int]$policy.BulkThreshold } else { $null }
            SpamAction = if ($policy.SpamAction) { $policy.SpamAction.ToString() } else { $null }
            HighConfidenceSpamAction = if ($policy.HighConfidenceSpamAction) { $policy.HighConfidenceSpamAction.ToString() } else { $null }
            PhishSpamAction = if ($policy.PhishSpamAction) { $policy.PhishSpamAction.ToString() } else { $null }
        }
    }";

        var results = await RunScriptAsync(script, cancellationToken);
        return results.Select(obj => new HostedContentFilterPolicyDto
        {
            Identity = GetString(obj, "Identity"),
            Name = GetString(obj, "Name"),
            RuleIdentity = GetNullableString(obj, "RuleIdentity"),
            RuleName = GetNullableString(obj, "RuleName"),
            RuleState = GetNullableString(obj, "RuleState"),
            Priority = GetNullableInt(obj, "Priority"),
            BulkThreshold = GetNullableInt(obj, "BulkThreshold"),
            SpamAction = GetNullableString(obj, "SpamAction"),
            HighConfidenceSpamAction = GetNullableString(obj, "HighConfidenceSpamAction"),
            PhishSpamAction = GetNullableString(obj, "PhishSpamAction")
        }).ToList();
    }

    private async Task<List<AntiPhishPolicyDto>> LoadAntiPhishPoliciesAsync(CancellationToken cancellationToken)
    {
        var script = @"
$rules = @()
if (Get-Command -Name Get-AntiPhishRule -ErrorAction SilentlyContinue) {
    $rules = @(Get-AntiPhishRule -ErrorAction Stop)
}

Get-AntiPhishPolicy -ErrorAction Stop |
    Sort-Object Name |
    ForEach-Object {
        $policy = $_
        $rule = $rules |
            Where-Object { $_.AntiPhishPolicy -eq $policy.Name } |
            Sort-Object Priority |
            Select-Object -First 1

        [PSCustomObject]@{
            Identity = $policy.Identity.ToString()
            Name = $policy.Name
            RuleIdentity = if ($rule) { $rule.Identity.ToString() } else { $null }
            RuleName = if ($rule) { $rule.Name } else { $null }
            RuleState = if ($rule) { $rule.State.ToString() } else { $null }
            Priority = if ($rule -and $rule.Priority -ne $null) { [int]$rule.Priority } else { $null }
            Enabled = if ($policy.Enabled -ne $null) { [bool]$policy.Enabled } else { $null }
            EnableSpoofIntelligence = if ($policy.EnableSpoofIntelligence -ne $null) { [bool]$policy.EnableSpoofIntelligence } else { $null }
            EnableMailboxIntelligence = if ($policy.EnableMailboxIntelligence -ne $null) { [bool]$policy.EnableMailboxIntelligence } else { $null }
            EnableTargetedUserProtection = if ($policy.EnableTargetedUserProtection -ne $null) { [bool]$policy.EnableTargetedUserProtection } else { $null }
            PhishThresholdLevel = if ($policy.PhishThresholdLevel -ne $null) { [int]$policy.PhishThresholdLevel } else { $null }
            MailboxIntelligenceProtectionAction = if ($policy.MailboxIntelligenceProtectionAction) { $policy.MailboxIntelligenceProtectionAction.ToString() } else { $null }
            TargetedUserProtectionAction = if ($policy.TargetedUserProtectionAction) { $policy.TargetedUserProtectionAction.ToString() } else { $null }
            AuthenticationFailAction = if ($policy.AuthenticationFailAction) { $policy.AuthenticationFailAction.ToString() } else { $null }
            HonorDmarcPolicy = if ($policy.HonorDmarcPolicy -ne $null) { [bool]$policy.HonorDmarcPolicy } else { $null }
            DmarcRejectAction = if ($policy.DmarcRejectAction) { $policy.DmarcRejectAction.ToString() } else { $null }
            DmarcQuarantineAction = if ($policy.DmarcQuarantineAction) { $policy.DmarcQuarantineAction.ToString() } else { $null }
        }
    }";

        var results = await RunScriptAsync(script, cancellationToken);
        return results.Select(obj => new AntiPhishPolicyDto
        {
            Identity = GetString(obj, "Identity"),
            Name = GetString(obj, "Name"),
            RuleIdentity = GetNullableString(obj, "RuleIdentity"),
            RuleName = GetNullableString(obj, "RuleName"),
            RuleState = GetNullableString(obj, "RuleState"),
            Priority = GetNullableInt(obj, "Priority"),
            Enabled = GetNullableBool(obj, "Enabled"),
            EnableSpoofIntelligence = GetNullableBool(obj, "EnableSpoofIntelligence"),
            EnableMailboxIntelligence = GetNullableBool(obj, "EnableMailboxIntelligence"),
            EnableTargetedUserProtection = GetNullableBool(obj, "EnableTargetedUserProtection"),
            PhishThresholdLevel = GetNullableInt(obj, "PhishThresholdLevel"),
            MailboxIntelligenceProtectionAction = GetNullableString(obj, "MailboxIntelligenceProtectionAction"),
            TargetedUserProtectionAction = GetNullableString(obj, "TargetedUserProtectionAction"),
            AuthenticationFailAction = GetNullableString(obj, "AuthenticationFailAction"),
            HonorDmarcPolicy = GetNullableBool(obj, "HonorDmarcPolicy"),
            DmarcRejectAction = GetNullableString(obj, "DmarcRejectAction"),
            DmarcQuarantineAction = GetNullableString(obj, "DmarcQuarantineAction")
        }).ToList();
    }

    private async Task<List<MalwareFilterPolicyDto>> LoadMalwareFilterPoliciesAsync(CancellationToken cancellationToken)
    {
        var script = @"
$rules = @()
if (Get-Command -Name Get-MalwareFilterRule -ErrorAction SilentlyContinue) {
    $rules = @(Get-MalwareFilterRule -ErrorAction Stop)
}

Get-MalwareFilterPolicy -ErrorAction Stop |
    Sort-Object Name |
    ForEach-Object {
        $policy = $_
        $rule = $rules |
            Where-Object { $_.MalwareFilterPolicy -eq $policy.Name } |
            Sort-Object Priority |
            Select-Object -First 1

        [PSCustomObject]@{
            Identity = $policy.Identity.ToString()
            Name = $policy.Name
            RuleIdentity = if ($rule) { $rule.Identity.ToString() } else { $null }
            RuleName = if ($rule) { $rule.Name } else { $null }
            RuleState = if ($rule) { $rule.State.ToString() } else { $null }
            Priority = if ($rule -and $rule.Priority -ne $null) { [int]$rule.Priority } else { $null }
            EnableFileFilter = if ($policy.EnableFileFilter -ne $null) { [bool]$policy.EnableFileFilter } else { $null }
            FileTypeAction = if ($policy.FileTypeAction) { $policy.FileTypeAction.ToString() } else { $null }
            ZapEnabled = if ($policy.ZapEnabled -ne $null) { [bool]$policy.ZapEnabled } else { $null }
        }
    }";

        var results = await RunScriptAsync(script, cancellationToken);
        return results.Select(obj => new MalwareFilterPolicyDto
        {
            Identity = GetString(obj, "Identity"),
            Name = GetString(obj, "Name"),
            RuleIdentity = GetNullableString(obj, "RuleIdentity"),
            RuleName = GetNullableString(obj, "RuleName"),
            RuleState = GetNullableString(obj, "RuleState"),
            Priority = GetNullableInt(obj, "Priority"),
            EnableFileFilter = GetNullableBool(obj, "EnableFileFilter"),
            FileTypeAction = GetNullableString(obj, "FileTypeAction"),
            ZapEnabled = GetNullableBool(obj, "ZapEnabled")
        }).ToList();
    }

    private async Task<List<QuarantinePolicyDto>> LoadQuarantinePoliciesAsync(CancellationToken cancellationToken)
    {
        var script = @"
Get-QuarantinePolicy -ErrorAction Stop |
    Sort-Object Name |
    ForEach-Object {
        [PSCustomObject]@{
            Identity = $_.Identity.ToString()
            Name = $_.Name
            EndUserQuarantinePermissionsValue = if ($_.EndUserQuarantinePermissionsValue) { $_.EndUserQuarantinePermissionsValue.ToString() } else { $null }
            ESNEnabled = if ($_.ESNEnabled -ne $null) { [bool]$_.ESNEnabled } else { $null }
            EndUserSpamNotificationFrequency = if ($_.EndUserSpamNotificationFrequency -ne $null) { [int]$_.EndUserSpamNotificationFrequency } else { $null }
            QuarantineRetentionDays = if ($_.QuarantineRetentionDays -ne $null) { [int]$_.QuarantineRetentionDays } else { $null }
            OrganizationBrandingEnabled = if ($_.OrganizationBrandingEnabled -ne $null) { [bool]$_.OrganizationBrandingEnabled } else { $null }
        }
    }";

        var results = await RunScriptAsync(script, cancellationToken);
        return results.Select(obj => new QuarantinePolicyDto
        {
            Identity = GetString(obj, "Identity"),
            Name = GetString(obj, "Name"),
            EndUserQuarantinePermissionsValue = GetNullableString(obj, "EndUserQuarantinePermissionsValue"),
            EsnEnabled = GetNullableBool(obj, "ESNEnabled"),
            EndUserSpamNotificationFrequency = GetNullableInt(obj, "EndUserSpamNotificationFrequency"),
            QuarantineRetentionDays = GetNullableInt(obj, "QuarantineRetentionDays"),
            OrganizationBrandingEnabled = GetNullableBool(obj, "OrganizationBrandingEnabled")
        }).ToList();
    }

    private async Task<List<HostedOutboundSpamFilterPolicyDto>> LoadHostedOutboundSpamFilterPoliciesAsync(CancellationToken cancellationToken)
    {
        var script = @"
Get-HostedOutboundSpamFilterPolicy -ErrorAction Stop |
    Sort-Object Name |
    ForEach-Object {
        [PSCustomObject]@{
            Identity = $_.Identity.ToString()
            Name = $_.Name
            RecipientLimitExternalPerHour = if ($_.RecipientLimitExternalPerHour -ne $null) { [int]$_.RecipientLimitExternalPerHour } else { $null }
            RecipientLimitInternalPerHour = if ($_.RecipientLimitInternalPerHour -ne $null) { [int]$_.RecipientLimitInternalPerHour } else { $null }
            RecipientLimitPerDay = if ($_.RecipientLimitPerDay -ne $null) { [int]$_.RecipientLimitPerDay } else { $null }
            ActionWhenThresholdReached = if ($_.ActionWhenThresholdReached) { $_.ActionWhenThresholdReached.ToString() } else { $null }
            AutoForwardingMode = if ($_.AutoForwardingMode) { $_.AutoForwardingMode.ToString() } else { $null }
        }
    }";

        var results = await RunScriptAsync(script, cancellationToken);
        return results.Select(obj => new HostedOutboundSpamFilterPolicyDto
        {
            Identity = GetString(obj, "Identity"),
            Name = GetString(obj, "Name"),
            RecipientLimitExternalPerHour = GetNullableInt(obj, "RecipientLimitExternalPerHour"),
            RecipientLimitInternalPerHour = GetNullableInt(obj, "RecipientLimitInternalPerHour"),
            RecipientLimitPerDay = GetNullableInt(obj, "RecipientLimitPerDay"),
            ActionWhenThresholdReached = GetNullableString(obj, "ActionWhenThresholdReached"),
            AutoForwardingMode = GetNullableString(obj, "AutoForwardingMode")
        }).ToList();
    }
}
