using System.Text.Json.Serialization;

namespace OnlyExo365.Contracts.Dtos;

public class GetTransportRulesRequest
{
}

public class GetTransportRulesResponse
{
    [JsonPropertyName("rules")]
    public List<TransportRuleDto> Rules { get; set; } = new();
}

public class TransportRuleDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("priority")]
    public int? Priority { get; set; }

    [JsonPropertyName("state")]
    public string State { get; set; } = string.Empty;

    [JsonPropertyName("mode")]
    public string Mode { get; set; } = string.Empty;

    [JsonPropertyName("description")]
    public string Description { get; set; } = string.Empty;

    [JsonPropertyName("from")]
    public List<string> From { get; set; } = new();

    [JsonPropertyName("sentTo")]
    public List<string> SentTo { get; set; } = new();

    [JsonPropertyName("senderDomainIs")]
    public List<string> SenderDomainIs { get; set; } = new();

    [JsonPropertyName("recipientDomainIs")]
    public List<string> RecipientDomainIs { get; set; } = new();

    [JsonPropertyName("sentToMemberOf")]
    public List<string> SentToMemberOf { get; set; } = new();

    [JsonPropertyName("subjectContainsWords")]
    public List<string> SubjectContainsWords { get; set; } = new();

    [JsonPropertyName("exceptIfFrom")]
    public List<string> ExceptIfFrom { get; set; } = new();

    [JsonPropertyName("exceptIfSentTo")]
    public List<string> ExceptIfSentTo { get; set; } = new();

    [JsonPropertyName("exceptIfSenderDomainIs")]
    public List<string> ExceptIfSenderDomainIs { get; set; } = new();

    [JsonPropertyName("exceptIfRecipientDomainIs")]
    public List<string> ExceptIfRecipientDomainIs { get; set; } = new();

    [JsonPropertyName("exceptIfSubjectContainsWords")]
    public List<string> ExceptIfSubjectContainsWords { get; set; } = new();

    [JsonPropertyName("prependSubject")]
    public string PrependSubject { get; set; } = string.Empty;

    [JsonPropertyName("redirectMessageTo")]
    public List<string> RedirectMessageTo { get; set; } = new();

    [JsonPropertyName("blindCopyTo")]
    public List<string> BlindCopyTo { get; set; } = new();

    [JsonPropertyName("addToRecipients")]
    public List<string> AddToRecipients { get; set; } = new();

    [JsonPropertyName("stopRuleProcessing")]
    public bool StopRuleProcessing { get; set; }

    [JsonPropertyName("deleteMessage")]
    public bool DeleteMessage { get; set; }
}

public class SetTransportRuleStateRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; }
}

public class RemoveTransportRuleRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}

public class UpsertTransportRuleRequest
{
    [JsonPropertyName("identity")]
    public string? Identity { get; set; }

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("from")]
    public List<string> From { get; set; } = new();

    [JsonPropertyName("sentTo")]
    public List<string> SentTo { get; set; } = new();

    [JsonPropertyName("senderDomainIs")]
    public List<string> SenderDomainIs { get; set; } = new();

    [JsonPropertyName("recipientDomainIs")]
    public List<string> RecipientDomainIs { get; set; } = new();

    [JsonPropertyName("sentToMemberOf")]
    public List<string> SentToMemberOf { get; set; } = new();

    [JsonPropertyName("subjectContainsWords")]
    public List<string> SubjectContainsWords { get; set; } = new();

    [JsonPropertyName("exceptIfFrom")]
    public List<string> ExceptIfFrom { get; set; } = new();

    [JsonPropertyName("exceptIfSentTo")]
    public List<string> ExceptIfSentTo { get; set; } = new();

    [JsonPropertyName("exceptIfSenderDomainIs")]
    public List<string> ExceptIfSenderDomainIs { get; set; } = new();

    [JsonPropertyName("exceptIfRecipientDomainIs")]
    public List<string> ExceptIfRecipientDomainIs { get; set; } = new();

    [JsonPropertyName("exceptIfSubjectContainsWords")]
    public List<string> ExceptIfSubjectContainsWords { get; set; } = new();

    [JsonPropertyName("prependSubject")]
    public string? PrependSubject { get; set; }

    [JsonPropertyName("redirectMessageTo")]
    public List<string> RedirectMessageTo { get; set; } = new();

    [JsonPropertyName("blindCopyTo")]
    public List<string> BlindCopyTo { get; set; } = new();

    [JsonPropertyName("addToRecipients")]
    public List<string> AddToRecipients { get; set; } = new();

    [JsonPropertyName("stopRuleProcessing")]
    public bool StopRuleProcessing { get; set; }

    [JsonPropertyName("deleteMessage")]
    public bool DeleteMessage { get; set; }

    [JsonPropertyName("mode")]
    public string? Mode { get; set; }

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; } = true;
}

public class TestTransportRuleRequest
{
    [JsonPropertyName("sender")]
    public string Sender { get; set; } = string.Empty;

    [JsonPropertyName("recipient")]
    public string Recipient { get; set; } = string.Empty;

    [JsonPropertyName("subject")]
    public string Subject { get; set; } = string.Empty;
}

public class TestTransportRuleResponse
{
    [JsonPropertyName("matchedRuleNames")]
    public List<string> MatchedRuleNames { get; set; } = new();
}

public class GetConnectorsRequest
{
}

public class GetConnectorsResponse
{
    [JsonPropertyName("connectors")]
    public List<ConnectorDto> Connectors { get; set; } = new();
}

public class ConnectorDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("displayLabel")]
    public string DisplayLabel { get; set; } = string.Empty;

    [JsonPropertyName("type")]
    public string Type { get; set; } = string.Empty;

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; }

    [JsonPropertyName("comment")]
    public string Comment { get; set; } = string.Empty;

    [JsonPropertyName("senderDomains")]
    public List<string> SenderDomains { get; set; } = new();

    [JsonPropertyName("recipientDomains")]
    public List<string> RecipientDomains { get; set; } = new();
}

public class UpsertConnectorRequest
{
    [JsonPropertyName("identity")]
    public string? Identity { get; set; }

    [JsonPropertyName("type")]
    public string Type { get; set; } = "Inbound";

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; } = true;

    [JsonPropertyName("comment")]
    public string? Comment { get; set; }

    [JsonPropertyName("senderDomains")]
    public List<string> SenderDomains { get; set; } = new();

    [JsonPropertyName("recipientDomains")]
    public List<string> RecipientDomains { get; set; } = new();
}

public class RemoveConnectorRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("type")]
    public string Type { get; set; } = string.Empty;
}

public class GetAcceptedDomainsRequest
{
}

public class GetAcceptedDomainsResponse
{
    [JsonPropertyName("domains")]
    public List<AcceptedDomainDto> Domains { get; set; } = new();
}

public class AcceptedDomainDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("domainName")]
    public string DomainName { get; set; } = string.Empty;

    [JsonPropertyName("domainType")]
    public string DomainType { get; set; } = string.Empty;

    [JsonPropertyName("default")]
    public bool Default { get; set; }
}

public class UpsertAcceptedDomainRequest
{
    [JsonPropertyName("identity")]
    public string? Identity { get; set; }

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("domainName")]
    public string DomainName { get; set; } = string.Empty;

    [JsonPropertyName("domainType")]
    public string DomainType { get; set; } = "Authoritative";

    [JsonPropertyName("makeDefault")]
    public bool MakeDefault { get; set; }
}

public class RemoveAcceptedDomainRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}

public class GetRemoteDomainsRequest
{
}

public class GetRemoteDomainsResponse
{
    [JsonPropertyName("domains")]
    public List<RemoteDomainDto> Domains { get; set; } = new();
}

public class RemoteDomainDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("domainName")]
    public string DomainName { get; set; } = string.Empty;

    [JsonPropertyName("allowedOofType")]
    public string AllowedOOFType { get; set; } = string.Empty;

    [JsonPropertyName("autoReplyEnabled")]
    public bool AutoReplyEnabled { get; set; }

    [JsonPropertyName("autoForwardEnabled")]
    public bool AutoForwardEnabled { get; set; }

    [JsonPropertyName("deliveryReportEnabled")]
    public bool DeliveryReportEnabled { get; set; }

    [JsonPropertyName("ndrEnabled")]
    public bool NDREnabled { get; set; }

    [JsonPropertyName("meetingForwardNotificationEnabled")]
    public bool MeetingForwardNotificationEnabled { get; set; }

    [JsonPropertyName("tnefEnabled")]
    public bool TNEFEnabled { get; set; }

    [JsonPropertyName("trustedMailOutboundEnabled")]
    public bool TrustedMailOutboundEnabled { get; set; }

    [JsonPropertyName("isDefault")]
    public bool IsDefault { get; set; }
}

public class UpsertRemoteDomainRequest
{
    [JsonPropertyName("identity")]
    public string? Identity { get; set; }

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("domainName")]
    public string DomainName { get; set; } = string.Empty;

    [JsonPropertyName("allowedOofType")]
    public string AllowedOOFType { get; set; } = "External";

    [JsonPropertyName("autoReplyEnabled")]
    public bool AutoReplyEnabled { get; set; } = true;

    [JsonPropertyName("autoForwardEnabled")]
    public bool AutoForwardEnabled { get; set; } = true;

    [JsonPropertyName("deliveryReportEnabled")]
    public bool DeliveryReportEnabled { get; set; } = true;

    [JsonPropertyName("ndrEnabled")]
    public bool NDREnabled { get; set; } = true;

    [JsonPropertyName("meetingForwardNotificationEnabled")]
    public bool MeetingForwardNotificationEnabled { get; set; } = true;

    [JsonPropertyName("tnefEnabled")]
    public bool TNEFEnabled { get; set; }

    [JsonPropertyName("trustedMailOutboundEnabled")]
    public bool TrustedMailOutboundEnabled { get; set; }
}

public class RemoveRemoteDomainRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}

public class GetOrganizationRelationshipsRequest
{
}

public class GetOrganizationRelationshipsResponse
{
    [JsonPropertyName("relationships")]
    public List<OrganizationRelationshipDto> Relationships { get; set; } = new();
}

public class OrganizationRelationshipDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("domainNames")]
    public List<string> DomainNames { get; set; } = new();

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; }

    [JsonPropertyName("freeBusyAccessEnabled")]
    public bool FreeBusyAccessEnabled { get; set; }

    [JsonPropertyName("freeBusyAccessLevel")]
    public string FreeBusyAccessLevel { get; set; } = string.Empty;

    [JsonPropertyName("mailTipsAccessEnabled")]
    public bool MailTipsAccessEnabled { get; set; }

    [JsonPropertyName("mailTipsAccessLevel")]
    public string MailTipsAccessLevel { get; set; } = string.Empty;

    [JsonPropertyName("targetApplicationUri")]
    public string? TargetApplicationUri { get; set; }

    [JsonPropertyName("targetAutodiscoverEpr")]
    public string? TargetAutodiscoverEpr { get; set; }

    [JsonPropertyName("archiveAccessEnabled")]
    public bool? ArchiveAccessEnabled { get; set; }

    [JsonPropertyName("deliveryReportEnabled")]
    public bool? DeliveryReportEnabled { get; set; }

    [JsonPropertyName("mailboxMoveEnabled")]
    public bool? MailboxMoveEnabled { get; set; }

    [JsonPropertyName("photosEnabled")]
    public bool? PhotosEnabled { get; set; }
}

public class UpsertOrganizationRelationshipRequest
{
    [JsonPropertyName("identity")]
    public string? Identity { get; set; }

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("domainNames")]
    public List<string> DomainNames { get; set; } = new();

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; } = true;

    [JsonPropertyName("freeBusyAccessEnabled")]
    public bool FreeBusyAccessEnabled { get; set; } = true;

    [JsonPropertyName("freeBusyAccessLevel")]
    public string FreeBusyAccessLevel { get; set; } = "AvailabilityOnly";

    [JsonPropertyName("mailTipsAccessEnabled")]
    public bool MailTipsAccessEnabled { get; set; }

    [JsonPropertyName("mailTipsAccessLevel")]
    public string MailTipsAccessLevel { get; set; } = "All";

    [JsonPropertyName("targetApplicationUri")]
    public string? TargetApplicationUri { get; set; }

    [JsonPropertyName("targetAutodiscoverEpr")]
    public string? TargetAutodiscoverEpr { get; set; }

    [JsonPropertyName("archiveAccessEnabled")]
    public bool? ArchiveAccessEnabled { get; set; }

    [JsonPropertyName("deliveryReportEnabled")]
    public bool? DeliveryReportEnabled { get; set; }

    [JsonPropertyName("mailboxMoveEnabled")]
    public bool? MailboxMoveEnabled { get; set; }

    [JsonPropertyName("photosEnabled")]
    public bool? PhotosEnabled { get; set; }
}

public class RemoveOrganizationRelationshipRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}

public class GetAddressListsRequest
{
}

public class GetAddressListsResponse
{
    [JsonPropertyName("addressLists")]
    public List<AddressListDto> AddressLists { get; set; } = new();

    [JsonPropertyName("warnings")]
    public List<string> Warnings { get; set; } = new();

    [JsonPropertyName("warningDetails")]
    public List<OperationWarningDto> WarningDetails { get; set; } = new();

    [JsonPropertyName("hasPartialData")]
    public bool HasPartialData { get; set; }

    [JsonPropertyName("isUnsupported")]
    public bool IsUnsupported { get; set; }
}

public class AddressListDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("recipientFilter")]
    public string RecipientFilter { get; set; } = string.Empty;

    [JsonPropertyName("recipientContainer")]
    public string? RecipientContainer { get; set; }

    [JsonPropertyName("includedRecipients")]
    public List<string> IncludedRecipients { get; set; } = new();

    [JsonPropertyName("conditionalCompany")]
    public List<string> ConditionalCompany { get; set; } = new();

    [JsonPropertyName("conditionalDepartment")]
    public List<string> ConditionalDepartment { get; set; } = new();

    [JsonPropertyName("conditionalStateOrProvince")]
    public List<string> ConditionalStateOrProvince { get; set; } = new();

    [JsonPropertyName("conditionalCustomAttribute1")]
    public List<string> ConditionalCustomAttribute1 { get; set; } = new();
}

public class UpsertAddressListRequest
{
    [JsonPropertyName("identity")]
    public string? Identity { get; set; }

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("displayName")]
    public string? DisplayName { get; set; }

    [JsonPropertyName("recipientFilter")]
    public string? RecipientFilter { get; set; }

    [JsonPropertyName("recipientContainer")]
    public string? RecipientContainer { get; set; }

    [JsonPropertyName("includedRecipients")]
    public List<string> IncludedRecipients { get; set; } = new();

    [JsonPropertyName("conditionalCompany")]
    public List<string> ConditionalCompany { get; set; } = new();

    [JsonPropertyName("conditionalDepartment")]
    public List<string> ConditionalDepartment { get; set; } = new();

    [JsonPropertyName("conditionalStateOrProvince")]
    public List<string> ConditionalStateOrProvince { get; set; } = new();

    [JsonPropertyName("conditionalCustomAttribute1")]
    public List<string> ConditionalCustomAttribute1 { get; set; } = new();
}

public class RemoveAddressListRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}

public class GetAddressBookPoliciesRequest
{
}

public class GetAddressBookPoliciesResponse
{
    [JsonPropertyName("policies")]
    public List<AddressBookPolicyDto> Policies { get; set; } = new();
}

public class AddressBookPolicyDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("addressLists")]
    public List<string> AddressLists { get; set; } = new();

    [JsonPropertyName("globalAddressList")]
    public string GlobalAddressList { get; set; } = string.Empty;

    [JsonPropertyName("offlineAddressBook")]
    public string OfflineAddressBook { get; set; } = string.Empty;

    [JsonPropertyName("roomList")]
    public string RoomList { get; set; } = string.Empty;
}

public class UpsertAddressBookPolicyRequest
{
    [JsonPropertyName("identity")]
    public string? Identity { get; set; }

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("addressLists")]
    public List<string> AddressLists { get; set; } = new();

    [JsonPropertyName("globalAddressList")]
    public string GlobalAddressList { get; set; } = string.Empty;

    [JsonPropertyName("offlineAddressBook")]
    public string OfflineAddressBook { get; set; } = string.Empty;

    [JsonPropertyName("roomList")]
    public string RoomList { get; set; } = string.Empty;
}

public class RemoveAddressBookPolicyRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}

public class GetOfflineAddressBooksRequest
{
}

public class GetOfflineAddressBooksResponse
{
    [JsonPropertyName("offlineAddressBooks")]
    public List<OfflineAddressBookDto> OfflineAddressBooks { get; set; } = new();

    [JsonPropertyName("warnings")]
    public List<string> Warnings { get; set; } = new();

    [JsonPropertyName("warningDetails")]
    public List<OperationWarningDto> WarningDetails { get; set; } = new();

    [JsonPropertyName("hasPartialData")]
    public bool HasPartialData { get; set; }

    [JsonPropertyName("isUnsupported")]
    public bool IsUnsupported { get; set; }
}

public class OfflineAddressBookDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("addressLists")]
    public List<string> AddressLists { get; set; } = new();

    [JsonPropertyName("diffRetentionPeriod")]
    public int? DiffRetentionPeriod { get; set; }

    [JsonPropertyName("isDefault")]
    public bool IsDefault { get; set; }
}

public class UpsertOfflineAddressBookRequest
{
    [JsonPropertyName("identity")]
    public string? Identity { get; set; }

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("addressLists")]
    public List<string> AddressLists { get; set; } = new();

    [JsonPropertyName("diffRetentionPeriod")]
    public int? DiffRetentionPeriod { get; set; }
}

public class RemoveOfflineAddressBookRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}

public class GetSharingPoliciesRequest
{
}

public class GetSharingPoliciesResponse
{
    [JsonPropertyName("policies")]
    public List<SharingPolicyDto> Policies { get; set; } = new();
}

public class SharingPolicyDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("domains")]
    public List<string> Domains { get; set; } = new();

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; }

    [JsonPropertyName("isDefault")]
    public bool IsDefault { get; set; }
}

public class UpsertSharingPolicyRequest
{
    [JsonPropertyName("identity")]
    public string? Identity { get; set; }

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("domains")]
    public List<string> Domains { get; set; } = new();

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; } = true;

    [JsonPropertyName("makeDefault")]
    public bool MakeDefault { get; set; }
}

public class RemoveSharingPolicyRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}

