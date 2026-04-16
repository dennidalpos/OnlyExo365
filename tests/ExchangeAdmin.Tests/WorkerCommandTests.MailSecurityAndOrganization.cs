using ExchangeAdmin.Contracts;
using System.Text.Json;
using ExchangeAdmin.Contracts.Diagnostics;
using ExchangeAdmin.Worker;
using ExchangeAdmin.Worker.PowerShell;

namespace ExchangeAdmin.Tests;

public partial class WorkerCommandTests
{
    [Fact]
    public void ConsoleLogger_SanitizesSensitiveDataBeforePersisting()
    {
        var tempDirectory = Path.Combine(Path.GetTempPath(), "ExchangeAdmin.Tests", Guid.NewGuid().ToString("N"));
        var correlationId = Guid.NewGuid().ToString("N");
        Directory.CreateDirectory(tempDirectory);

        try
        {
            var writer = new PersistentLogWriter(
                "worker",
                tempDirectory,
                utcNow: () => new DateTime(2026, 3, 11, 8, 0, 0, DateTimeKind.Utc),
                processId: () => 31337);

            ConsoleLogger.SetPersistentLogWriterForTesting(writer);
            using (ConsoleLogger.BeginCorrelationScope(correlationId))
            {
                ConsoleLogger.Verbose("PowerShellEngine", "Script preview (sanitized): $securePassword = ConvertTo-SecureString 'Sup3rSecret!' -AsPlainText -Force");
            }
            ConsoleLogger.ResetPersistentLogWriterForTesting();

            var filePath = Path.Combine(tempDirectory, "worker-20260311.log");
            var line = File.ReadAllLines(filePath)
                .Single(entry => entry.Contains(correlationId, StringComparison.Ordinal));
            var payload = JsonSerializer.Deserialize<PersistentLogEntry>(line);

            Assert.DoesNotContain("Sup3rSecret!", line, StringComparison.Ordinal);
            Assert.NotNull(payload);
            Assert.Equal(correlationId, payload!.CorrelationId);
            Assert.Contains("<redacted>", payload!.Message, StringComparison.Ordinal);
        }
        finally
        {
            ConsoleLogger.ResetPersistentLogWriterForTesting();
            if (Directory.Exists(tempDirectory))
            {
                Directory.Delete(tempDirectory, recursive: true);
            }
        }
    }

    [Fact]
    public void BuildUpsertTransportRuleScript_IncludesAdvancedConditionsExceptionsAndActions()
    {
        var request = new UpsertTransportRuleRequest
        {
            Identity = "Rule'A",
            Name = "Block external finance",
            From = new List<string> { "ceo@contoso.com" },
            SenderDomainIs = new List<string> { "contoso.com" },
            RecipientDomainIs = new List<string> { "fabrikam.com" },
            SentToMemberOf = new List<string> { "Finance Reviewers" },
            SubjectContainsWords = new List<string> { "wire" },
            ExceptIfSenderDomainIs = new List<string> { "trusted.partner" },
            ExceptIfSubjectContainsWords = new List<string> { "approved" },
            RedirectMessageTo = new List<string> { "review@contoso.com" },
            BlindCopyTo = new List<string> { "audit@contoso.com" },
            AddToRecipients = new List<string> { "compliance@contoso.com" },
            PrependSubject = "[REVIEW]",
            StopRuleProcessing = true,
            DeleteMessage = true,
            Enabled = true,
            Mode = "Enforce"
        };

        var script = TransportRuleCommandBuilder.BuildUpsertTransportRuleScript(request);

        Assert.Contains("$params['SenderDomainIs'] = $senderDomainIs", script, StringComparison.Ordinal);
        Assert.Contains("$params['RecipientDomainIs'] = $recipientDomainIs", script, StringComparison.Ordinal);
        Assert.Contains("$params['SentToMemberOf'] = $sentToMemberOf", script, StringComparison.Ordinal);
        Assert.Contains("$params['ExceptIfSenderDomainIs'] = $exceptIfSenderDomainIs", script, StringComparison.Ordinal);
        Assert.Contains("$params['ExceptIfSubjectContainsWords'] = $exceptIfSubjectContains", script, StringComparison.Ordinal);
        Assert.Contains("$params['RedirectMessageTo'] = $redirectMessageTo", script, StringComparison.Ordinal);
        Assert.Contains("$params['BlindCopyTo'] = $blindCopyTo", script, StringComparison.Ordinal);
        Assert.Contains("$params['AddToRecipients'] = $addToRecipients", script, StringComparison.Ordinal);
        Assert.Contains("$params['StopRuleProcessing'] = $true", script, StringComparison.Ordinal);
        Assert.Contains("$params['DeleteMessage'] = $true", script, StringComparison.Ordinal);
        Assert.Contains("Set-TransportRule -Identity 'Rule''A' @params -ErrorAction Stop", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildTestTransportRuleScript_ChecksDomainsAndExceptions()
    {
        var request = new TestTransportRuleRequest
        {
            Sender = "sender@contoso.com",
            Recipient = "user@fabrikam.com",
            Subject = "wire approved"
        };

        var script = TransportRuleCommandBuilder.BuildTestTransportRuleScript(request);

        Assert.Contains("$senderDomain = if ($sender -like '*@*')", script, StringComparison.Ordinal);
        Assert.Contains("$recipientDomain = if ($recipient -like '*@*')", script, StringComparison.Ordinal);
        Assert.Contains("$ruleSenderDomains -contains $senderDomain", script, StringComparison.Ordinal);
        Assert.Contains("$ruleRecipientDomains -contains $recipientDomain", script, StringComparison.Ordinal);
        Assert.Contains("$ruleExceptSenderDomains -contains $senderDomain", script, StringComparison.Ordinal);
        Assert.Contains("$ruleExceptRecipientDomains -contains $recipientDomain", script, StringComparison.Ordinal);
        Assert.Contains("$r.ExceptIfSubjectContainsWords", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildUpsertRemoteDomainScript_UsesExplicitToggleParameters()
    {
        var request = new UpsertRemoteDomainRequest
        {
            Identity = "Partner Domain",
            Name = "Partner Domain",
            DomainName = "*.partner.example",
            AllowedOOFType = "ExternalLegacy",
            AutoReplyEnabled = true,
            AutoForwardEnabled = false,
            DeliveryReportEnabled = true,
            NDREnabled = false,
            MeetingForwardNotificationEnabled = true,
            TNEFEnabled = true,
            TrustedMailOutboundEnabled = false
        };

        var script = RemoteDomainCommandBuilder.BuildUpsertRemoteDomainScript(request);

        Assert.Contains("AllowedOOFType = 'ExternalLegacy'", script, StringComparison.Ordinal);
        Assert.Contains("AutoForwardEnabled = $false", script, StringComparison.Ordinal);
        Assert.Contains("NDREnabled = $false", script, StringComparison.Ordinal);
        Assert.Contains("TNEFEnabled = $true", script, StringComparison.Ordinal);
        Assert.Contains("$params['Name'] = 'Partner Domain'", script, StringComparison.Ordinal);
        Assert.Contains("Set-RemoteDomain -Identity 'Partner Domain' @params -ErrorAction Stop", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildUpsertOrganizationRelationshipScript_UsesDomainListAndCrossTenantFlags()
    {
        var request = new UpsertOrganizationRelationshipRequest
        {
            Identity = "CrossTenant",
            Name = "CrossTenant",
            DomainNames = new List<string> { "contoso.com", "fabrikam.com" },
            Enabled = true,
            FreeBusyAccessEnabled = true,
            FreeBusyAccessLevel = "LimitedDetails",
            MailTipsAccessEnabled = true,
            MailTipsAccessLevel = "Limited",
            TargetApplicationUri = "https://outlook.com/",
            TargetAutodiscoverEpr = "https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity",
            ArchiveAccessEnabled = false,
            DeliveryReportEnabled = true,
            MailboxMoveEnabled = false,
            PhotosEnabled = true
        };

        var script = OrganizationRelationshipCommandBuilder.BuildUpsertOrganizationRelationshipScript(request);

        Assert.Contains("$domainNames = @('contoso.com', 'fabrikam.com')", script, StringComparison.Ordinal);
        Assert.Contains("FreeBusyAccessLevel = 'LimitedDetails'", script, StringComparison.Ordinal);
        Assert.Contains("MailTipsAccessLevel = 'Limited'", script, StringComparison.Ordinal);
        Assert.Contains("$params['TargetApplicationUri'] = 'https://outlook.com/'", script, StringComparison.Ordinal);
        Assert.Contains("$params['TargetAutodiscoverEpr'] = 'https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity'", script, StringComparison.Ordinal);
        Assert.Contains("DeliveryReportEnabled = $true", script, StringComparison.Ordinal);
        Assert.Contains("PhotosEnabled = $true", script, StringComparison.Ordinal);
        Assert.Contains("Set-OrganizationRelationship -Identity 'CrossTenant' @params -ErrorAction Stop", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildUpsertAddressListScript_UsesFilterOrConditionalParametersSafely()
    {
        var request = new UpsertAddressListRequest
        {
            Identity = "\\All Staff",
            Name = "All Staff",
            DisplayName = "All Staff",
            IncludedRecipients = new List<string> { "MailboxUsers", "MailUsers" },
            ConditionalCompany = new List<string> { "Contoso" },
            ConditionalDepartment = new List<string> { "Finance" }
        };

        var script = OrganizationDirectoryCommandBuilder.BuildUpsertAddressListScript(request);

        Assert.Contains("$includedRecipients = @('MailboxUsers', 'MailUsers')", script, StringComparison.Ordinal);
        Assert.Contains("$conditionalCompany = @('Contoso')", script, StringComparison.Ordinal);
        Assert.Contains("$conditionalDepartment = @('Finance')", script, StringComparison.Ordinal);
        Assert.Contains("$params['IncludedRecipients'] = $includedRecipients", script, StringComparison.Ordinal);
        Assert.Contains("Set-AddressList -Identity '\\All Staff' @params -ErrorAction Stop", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildUpsertAddressBookPolicyScript_UsesGalOabAndOptionalRoomList()
    {
        var request = new UpsertAddressBookPolicyRequest
        {
            Name = "Executives",
            AddressLists = new List<string> { "\\Executives", "\\Rooms" },
            GlobalAddressList = "\\Default Global Address List",
            OfflineAddressBook = "\\Default Offline Address Book",
            RoomList = "\\Room Lists\\HQ"
        };

        var script = OrganizationDirectoryCommandBuilder.BuildUpsertAddressBookPolicyScript(request);

        Assert.Contains("$addressLists = @('\\Executives', '\\Rooms')", script, StringComparison.Ordinal);
        Assert.Contains("GlobalAddressList = '\\Default Global Address List'", script, StringComparison.Ordinal);
        Assert.Contains("OfflineAddressBook = '\\Default Offline Address Book'", script, StringComparison.Ordinal);
        Assert.Contains("$params['RoomList'] = '\\Room Lists\\HQ'", script, StringComparison.Ordinal);
        Assert.Contains("New-AddressBookPolicy @params -ErrorAction Stop", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildUpsertOfflineAddressBookScript_UsesAddressListsAndOptionalDiffRetention()
    {
        var request = new UpsertOfflineAddressBookRequest
        {
            Identity = "\\Default Offline Address Book",
            Name = "Default Offline Address Book",
            AddressLists = new List<string> { "\\Default Global Address List" },
            DiffRetentionPeriod = 30
        };

        var script = OrganizationDirectoryCommandBuilder.BuildUpsertOfflineAddressBookScript(request);

        Assert.Contains("$addressLists = @('\\Default Global Address List')", script, StringComparison.Ordinal);
        Assert.Contains("$params['DiffRetentionPeriod'] = 30", script, StringComparison.Ordinal);
        Assert.Contains("Set-OfflineAddressBook -Identity '\\Default Offline Address Book' @params -ErrorAction Stop", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildUpsertSharingPolicyScript_UsesDomainsEnabledAndDefaultPromotion()
    {
        var request = new UpsertSharingPolicyRequest
        {
            Identity = "Partner Sharing",
            Name = "Partner Sharing",
            Domains = new List<string> { "contoso.com: CalendarSharingFreeBusyDetail", "fabrikam.com: ContactsSharing" },
            Enabled = true,
            MakeDefault = true
        };

        var script = OrganizationDirectoryCommandBuilder.BuildUpsertSharingPolicyScript(request);

        Assert.Contains("$domains = @('contoso.com: CalendarSharingFreeBusyDetail', 'fabrikam.com: ContactsSharing')", script, StringComparison.Ordinal);
        Assert.Contains("Enabled = $true", script, StringComparison.Ordinal);
        Assert.Contains("Set-SharingPolicy -Identity 'Partner Sharing' @params -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("Set-SharingPolicy -Identity 'Partner Sharing' -Default:$true -ErrorAction Stop", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildUpdateHostedContentFilterPolicyScript_UpdatesPolicyAndLinkedRuleState()
    {
        var request = new UpdateHostedContentFilterPolicyRequest
        {
            Identity = "Default",
            RuleIdentity = "Default Rule",
            Enabled = false,
            BulkThreshold = 6,
            SpamAction = "MoveToJmf",
            HighConfidenceSpamAction = "Quarantine",
            PhishSpamAction = "Delete"
        };

        var script = ExoCommands.BuildUpdateHostedContentFilterPolicyScript(request);

        Assert.Contains("Set-HostedContentFilterPolicy -Identity 'Default' @params -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("$params['BulkThreshold'] = 6", script, StringComparison.Ordinal);
        Assert.Contains("$params['SpamAction'] = 'MoveToJmf'", script, StringComparison.Ordinal);
        Assert.Contains("$params['HighConfidenceSpamAction'] = 'Quarantine'", script, StringComparison.Ordinal);
        Assert.Contains("$params['PhishSpamAction'] = 'Delete'", script, StringComparison.Ordinal);
        Assert.Contains("Disable-HostedContentFilterRule -Identity 'Default Rule' -Confirm:$false -ErrorAction Stop", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildUpdateAntiPhishPolicyScript_UpdatesThresholdDmarcAndRuleState()
    {
        var request = new UpdateAntiPhishPolicyRequest
        {
            Identity = "Default AntiPhish",
            RuleIdentity = "Default AntiPhish Rule",
            Enabled = true,
            EnableSpoofIntelligence = true,
            EnableMailboxIntelligence = true,
            EnableTargetedUserProtection = false,
            HonorDmarcPolicy = true,
            PhishThresholdLevel = 3,
            MailboxIntelligenceProtectionAction = "Quarantine",
            AuthenticationFailAction = "MoveToJmf",
            DmarcRejectAction = "Quarantine",
            DmarcQuarantineAction = "Delete"
        };

        var script = ExoCommands.BuildUpdateAntiPhishPolicyScript(request);

        Assert.Contains("Set-AntiPhishPolicy -Identity 'Default AntiPhish' @params -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("$params['PhishThresholdLevel'] = 3", script, StringComparison.Ordinal);
        Assert.Contains("$params['EnableSpoofIntelligence'] = $true", script, StringComparison.Ordinal);
        Assert.Contains("$params['EnableMailboxIntelligence'] = $true", script, StringComparison.Ordinal);
        Assert.Contains("$params['EnableTargetedUserProtection'] = $false", script, StringComparison.Ordinal);
        Assert.Contains("$params['HonorDmarcPolicy'] = $true", script, StringComparison.Ordinal);
        Assert.Contains("$params['MailboxIntelligenceProtectionAction'] = 'Quarantine'", script, StringComparison.Ordinal);
        Assert.Contains("Enable-AntiPhishRule -Identity 'Default AntiPhish Rule' -Confirm:$false -ErrorAction Stop", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildUpdateHostedOutboundSpamFilterPolicyScript_UsesExplicitThresholdParameters()
    {
        var request = new UpdateHostedOutboundSpamFilterPolicyRequest
        {
            Identity = "Default Outbound",
            RecipientLimitExternalPerHour = 500,
            RecipientLimitInternalPerHour = 750,
            RecipientLimitPerDay = 1000,
            ActionWhenThresholdReached = "BlockUser",
            AutoForwardingMode = "Automatic"
        };

        var script = ExoCommands.BuildUpdateHostedOutboundSpamFilterPolicyScript(request);

        Assert.Contains("Set-HostedOutboundSpamFilterPolicy -Identity 'Default Outbound' @params -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("$params['RecipientLimitExternalPerHour'] = 500", script, StringComparison.Ordinal);
        Assert.Contains("$params['RecipientLimitInternalPerHour'] = 750", script, StringComparison.Ordinal);
        Assert.Contains("$params['RecipientLimitPerDay'] = 1000", script, StringComparison.Ordinal);
        Assert.Contains("$params['ActionWhenThresholdReached'] = 'BlockUser'", script, StringComparison.Ordinal);
        Assert.Contains("$params['AutoForwardingMode'] = 'Automatic'", script, StringComparison.Ordinal);
    }
}
