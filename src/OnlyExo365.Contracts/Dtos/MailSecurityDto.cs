using System.Text.Json.Serialization;

namespace OnlyExo365.Contracts.Dtos;

public class GetMailSecurityBaselineRequest
{
}

public class GetMailSecurityBaselineResponse
{
    [JsonPropertyName("dkimConfigs")]
    public List<DkimSigningConfigDto> DkimConfigs { get; set; } = new();

    [JsonPropertyName("antiSpamPolicies")]
    public List<HostedContentFilterPolicyDto> AntiSpamPolicies { get; set; } = new();

    [JsonPropertyName("antiPhishPolicies")]
    public List<AntiPhishPolicyDto> AntiPhishPolicies { get; set; } = new();

    [JsonPropertyName("malwarePolicies")]
    public List<MalwareFilterPolicyDto> MalwarePolicies { get; set; } = new();

    [JsonPropertyName("quarantinePolicies")]
    public List<QuarantinePolicyDto> QuarantinePolicies { get; set; } = new();

    [JsonPropertyName("outboundSpamPolicies")]
    public List<HostedOutboundSpamFilterPolicyDto> OutboundSpamPolicies { get; set; } = new();

    [JsonPropertyName("warnings")]
    public List<string> Warnings { get; set; } = new();
}

public class DkimSigningConfigDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("domain")]
    public string Domain { get; set; } = string.Empty;

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; }

    [JsonPropertyName("selector1CName")]
    public string? Selector1CName { get; set; }

    [JsonPropertyName("selector2CName")]
    public string? Selector2CName { get; set; }

    [JsonPropertyName("lastChecked")]
    public DateTime? LastChecked { get; set; }
}

public abstract class MailSecurityPolicyDtoBase
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;
}

public abstract class RuleLinkedMailSecurityPolicyDtoBase : MailSecurityPolicyDtoBase
{
    [JsonPropertyName("ruleIdentity")]
    public string? RuleIdentity { get; set; }

    [JsonPropertyName("ruleName")]
    public string? RuleName { get; set; }

    [JsonPropertyName("ruleState")]
    public string? RuleState { get; set; }

    [JsonPropertyName("priority")]
    public int? Priority { get; set; }
}

public class HostedContentFilterPolicyDto : RuleLinkedMailSecurityPolicyDtoBase
{
    [JsonPropertyName("bulkThreshold")]
    public int? BulkThreshold { get; set; }

    [JsonPropertyName("spamAction")]
    public string? SpamAction { get; set; }

    [JsonPropertyName("highConfidenceSpamAction")]
    public string? HighConfidenceSpamAction { get; set; }

    [JsonPropertyName("phishSpamAction")]
    public string? PhishSpamAction { get; set; }
}

public class AntiPhishPolicyDto : RuleLinkedMailSecurityPolicyDtoBase
{
    [JsonPropertyName("enabled")]
    public bool? Enabled { get; set; }

    [JsonPropertyName("enableSpoofIntelligence")]
    public bool? EnableSpoofIntelligence { get; set; }

    [JsonPropertyName("enableMailboxIntelligence")]
    public bool? EnableMailboxIntelligence { get; set; }

    [JsonPropertyName("enableTargetedUserProtection")]
    public bool? EnableTargetedUserProtection { get; set; }

    [JsonPropertyName("phishThresholdLevel")]
    public int? PhishThresholdLevel { get; set; }

    [JsonPropertyName("mailboxIntelligenceProtectionAction")]
    public string? MailboxIntelligenceProtectionAction { get; set; }

    [JsonPropertyName("targetedUserProtectionAction")]
    public string? TargetedUserProtectionAction { get; set; }

    [JsonPropertyName("authenticationFailAction")]
    public string? AuthenticationFailAction { get; set; }

    [JsonPropertyName("honorDmarcPolicy")]
    public bool? HonorDmarcPolicy { get; set; }

    [JsonPropertyName("dmarcRejectAction")]
    public string? DmarcRejectAction { get; set; }

    [JsonPropertyName("dmarcQuarantineAction")]
    public string? DmarcQuarantineAction { get; set; }
}

public class MalwareFilterPolicyDto : RuleLinkedMailSecurityPolicyDtoBase
{
    [JsonPropertyName("enableFileFilter")]
    public bool? EnableFileFilter { get; set; }

    [JsonPropertyName("fileTypeAction")]
    public string? FileTypeAction { get; set; }

    [JsonPropertyName("zapEnabled")]
    public bool? ZapEnabled { get; set; }
}

public class QuarantinePolicyDto : MailSecurityPolicyDtoBase
{
    [JsonPropertyName("endUserQuarantinePermissionsValue")]
    public string? EndUserQuarantinePermissionsValue { get; set; }

    [JsonPropertyName("esnEnabled")]
    public bool? EsnEnabled { get; set; }

    [JsonPropertyName("endUserSpamNotificationFrequency")]
    public int? EndUserSpamNotificationFrequency { get; set; }

    [JsonPropertyName("quarantineRetentionDays")]
    public int? QuarantineRetentionDays { get; set; }

    [JsonPropertyName("organizationBrandingEnabled")]
    public bool? OrganizationBrandingEnabled { get; set; }
}

public class HostedOutboundSpamFilterPolicyDto : MailSecurityPolicyDtoBase
{
    [JsonPropertyName("recipientLimitExternalPerHour")]
    public int? RecipientLimitExternalPerHour { get; set; }

    [JsonPropertyName("recipientLimitInternalPerHour")]
    public int? RecipientLimitInternalPerHour { get; set; }

    [JsonPropertyName("recipientLimitPerDay")]
    public int? RecipientLimitPerDay { get; set; }

    [JsonPropertyName("actionWhenThresholdReached")]
    public string? ActionWhenThresholdReached { get; set; }

    [JsonPropertyName("autoForwardingMode")]
    public string? AutoForwardingMode { get; set; }
}

public class UpdateDkimSigningConfigRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; }
}

public class UpdateHostedContentFilterPolicyRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("ruleIdentity")]
    public string? RuleIdentity { get; set; }

    [JsonPropertyName("enabled")]
    public bool? Enabled { get; set; }

    [JsonPropertyName("bulkThreshold")]
    public int? BulkThreshold { get; set; }

    [JsonPropertyName("spamAction")]
    public string? SpamAction { get; set; }

    [JsonPropertyName("highConfidenceSpamAction")]
    public string? HighConfidenceSpamAction { get; set; }

    [JsonPropertyName("phishSpamAction")]
    public string? PhishSpamAction { get; set; }
}

public class UpdateAntiPhishPolicyRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("ruleIdentity")]
    public string? RuleIdentity { get; set; }

    [JsonPropertyName("enabled")]
    public bool? Enabled { get; set; }

    [JsonPropertyName("enableSpoofIntelligence")]
    public bool? EnableSpoofIntelligence { get; set; }

    [JsonPropertyName("enableMailboxIntelligence")]
    public bool? EnableMailboxIntelligence { get; set; }

    [JsonPropertyName("enableTargetedUserProtection")]
    public bool? EnableTargetedUserProtection { get; set; }

    [JsonPropertyName("phishThresholdLevel")]
    public int? PhishThresholdLevel { get; set; }

    [JsonPropertyName("mailboxIntelligenceProtectionAction")]
    public string? MailboxIntelligenceProtectionAction { get; set; }

    [JsonPropertyName("targetedUserProtectionAction")]
    public string? TargetedUserProtectionAction { get; set; }

    [JsonPropertyName("authenticationFailAction")]
    public string? AuthenticationFailAction { get; set; }

    [JsonPropertyName("honorDmarcPolicy")]
    public bool? HonorDmarcPolicy { get; set; }

    [JsonPropertyName("dmarcRejectAction")]
    public string? DmarcRejectAction { get; set; }

    [JsonPropertyName("dmarcQuarantineAction")]
    public string? DmarcQuarantineAction { get; set; }
}

public class UpdateMalwareFilterPolicyRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("ruleIdentity")]
    public string? RuleIdentity { get; set; }

    [JsonPropertyName("enabled")]
    public bool? Enabled { get; set; }

    [JsonPropertyName("enableFileFilter")]
    public bool? EnableFileFilter { get; set; }

    [JsonPropertyName("fileTypeAction")]
    public string? FileTypeAction { get; set; }

    [JsonPropertyName("zapEnabled")]
    public bool? ZapEnabled { get; set; }
}

public class UpdateQuarantinePolicyRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("endUserQuarantinePermissionsValue")]
    public string? EndUserQuarantinePermissionsValue { get; set; }

    [JsonPropertyName("esnEnabled")]
    public bool? EsnEnabled { get; set; }

    [JsonPropertyName("endUserSpamNotificationFrequency")]
    public int? EndUserSpamNotificationFrequency { get; set; }

    [JsonPropertyName("quarantineRetentionDays")]
    public int? QuarantineRetentionDays { get; set; }

    [JsonPropertyName("organizationBrandingEnabled")]
    public bool? OrganizationBrandingEnabled { get; set; }
}

public class UpdateHostedOutboundSpamFilterPolicyRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("recipientLimitExternalPerHour")]
    public int? RecipientLimitExternalPerHour { get; set; }

    [JsonPropertyName("recipientLimitInternalPerHour")]
    public int? RecipientLimitInternalPerHour { get; set; }

    [JsonPropertyName("recipientLimitPerDay")]
    public int? RecipientLimitPerDay { get; set; }

    [JsonPropertyName("actionWhenThresholdReached")]
    public string? ActionWhenThresholdReached { get; set; }

    [JsonPropertyName("autoForwardingMode")]
    public string? AutoForwardingMode { get; set; }
}

