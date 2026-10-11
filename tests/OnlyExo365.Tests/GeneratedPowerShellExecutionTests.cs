using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Tests;

public class GeneratedPowerShellExecutionTests
{
    [Fact]
    public async Task TransportRuleUpdate_BindsEscapedIdentityAndConditions()
    {
        using var fixture = new LocalPowerShellFixture();
        await fixture.DefineCommandsAsync("""
            function Set-TransportRule {
                [CmdletBinding()]
                param(
                    [string]$Identity, [string]$Name, [string]$Mode, [bool]$Enabled,
                    [string[]]$SenderDomainIs, [string[]]$ExceptIfSenderDomainIs,
                    [string]$PrependSubject, [bool]$StopRuleProcessing
                )
                [pscustomobject]@{
                    Identity = $Identity
                    Name = $Name
                    SenderDomain = $SenderDomainIs[0]
                    ExceptDomain = $ExceptIfSenderDomainIs[0]
                    PrependSubject = $PrependSubject
                    StopRuleProcessing = $StopRuleProcessing
                }
            }
            function New-TransportRule { throw 'Unexpected create path' }
            """);

        var script = TransportRuleCommandBuilder.BuildUpsertTransportRuleScript(new UpsertTransportRuleRequest
        {
            Identity = "Rule'O'Brien",
            Name = "Finance rule",
            SenderDomainIs = ["contoso.com"],
            ExceptIfSenderDomainIs = ["trusted.example"],
            PrependSubject = "[REVIEW]",
            StopRuleProcessing = true,
            Enabled = true
        });

        var result = await fixture.Engine.ExecuteAsync(script);

        Assert.True(result.Success, result.ErrorMessage);
        var output = Assert.Single(result.Output);
        Assert.Equal("Rule'O'Brien", output.Properties["Identity"]?.Value);
        Assert.Equal("Finance rule", output.Properties["Name"]?.Value);
        Assert.Equal("contoso.com", output.Properties["SenderDomain"]?.Value);
        Assert.Equal("trusted.example", output.Properties["ExceptDomain"]?.Value);
        Assert.Equal("[REVIEW]", output.Properties["PrependSubject"]?.Value);
        Assert.Equal(true, output.Properties["StopRuleProcessing"]?.Value);
    }

    [Fact]
    public async Task TransportRuleUpdate_PropagatesCmdletFailure()
    {
        using var fixture = new LocalPowerShellFixture();
        await fixture.DefineCommandsAsync("""
            function Set-TransportRule {
                [CmdletBinding()]
                param([string]$Identity, [string]$Name, [string]$Mode, [bool]$Enabled)
                throw 'Offline write rejected'
            }
            """);

        var script = TransportRuleCommandBuilder.BuildUpsertTransportRuleScript(new UpsertTransportRuleRequest
        {
            Identity = "Finance",
            Name = "Finance",
            Enabled = true
        });

        var result = await fixture.Engine.ExecuteAsync(script);

        Assert.False(result.Success);
        Assert.Contains("Offline write rejected", result.ErrorMessage, StringComparison.Ordinal);
    }

    [Fact]
    public async Task RemoteDomainCreate_RunsCreateBeforeConfiguration()
    {
        using var fixture = new LocalPowerShellFixture();
        await fixture.DefineCommandsAsync("""
            function New-RemoteDomain {
                [CmdletBinding()]
                param([string]$Name, [string]$DomainName)
                $script:createdDomain = "$Name|$DomainName"
            }
            function Set-RemoteDomain {
                [CmdletBinding()]
                param(
                    [string]$Identity, [string]$AllowedOOFType, [bool]$AutoReplyEnabled,
                    [bool]$AutoForwardEnabled, [bool]$DeliveryReportEnabled, [bool]$NDREnabled,
                    [bool]$MeetingForwardNotificationEnabled, [bool]$TNEFEnabled,
                    [bool]$TrustedMailOutboundEnabled
                )
                if ($script:createdDomain -ne "$Identity|*.partner.example") { throw 'Create was not completed first' }
                [pscustomobject]@{ Identity = $Identity; AutoForwardEnabled = $AutoForwardEnabled; TNEFEnabled = $TNEFEnabled }
            }
            """);

        var script = RemoteDomainCommandBuilder.BuildUpsertRemoteDomainScript(new UpsertRemoteDomainRequest
        {
            Name = "Partner",
            DomainName = "*.partner.example",
            AutoForwardEnabled = false,
            TNEFEnabled = true
        });

        var result = await fixture.Engine.ExecuteAsync(script);

        Assert.True(result.Success, result.ErrorMessage);
        var output = Assert.Single(result.Output);
        Assert.Equal("Partner", output.Properties["Identity"]?.Value);
        Assert.Equal(false, output.Properties["AutoForwardEnabled"]?.Value);
        Assert.Equal(true, output.Properties["TNEFEnabled"]?.Value);
    }

    [Fact]
    public async Task AddressListUpdate_BindsConditionalFilters()
    {
        using var fixture = new LocalPowerShellFixture();
        await fixture.DefineCommandsAsync("""
            function Set-AddressList {
                [CmdletBinding()]
                param(
                    [string]$Identity, [string]$DisplayName, [string[]]$IncludedRecipients,
                    [string[]]$ConditionalCompany, [string[]]$ConditionalDepartment
                )
                [pscustomobject]@{
                    Identity = $Identity
                    IncludedRecipient = $IncludedRecipients[0]
                    Company = $ConditionalCompany[0]
                    Department = $ConditionalDepartment[0]
                    HasRecipientFilter = $PSBoundParameters.ContainsKey('RecipientFilter')
                }
            }
            function New-AddressList { throw 'Unexpected create path' }
            """);

        var script = OrganizationDirectoryCommandBuilder.BuildUpsertAddressListScript(new UpsertAddressListRequest
        {
            Identity = "\\All Staff",
            Name = "All Staff",
            IncludedRecipients = ["MailboxUsers"],
            ConditionalCompany = ["Contoso"],
            ConditionalDepartment = ["Finance"]
        });

        var result = await fixture.Engine.ExecuteAsync(script);

        Assert.True(result.Success, result.ErrorMessage);
        var output = Assert.Single(result.Output);
        Assert.Equal("\\All Staff", output.Properties["Identity"]?.Value);
        Assert.Equal("MailboxUsers", output.Properties["IncludedRecipient"]?.Value);
        Assert.Equal("Contoso", output.Properties["Company"]?.Value);
        Assert.Equal("Finance", output.Properties["Department"]?.Value);
        Assert.Equal(false, output.Properties["HasRecipientFilter"]?.Value);
    }

    [Fact]
    public async Task TransportRulePreview_AppliesExceptionToOtherwiseMatchingRule()
    {
        using var fixture = new LocalPowerShellFixture();
        await fixture.DefineCommandsAsync("""
            function Get-TransportRule {
                [CmdletBinding()]
                param()
                [pscustomobject]@{
                    Name = 'Allowed'
                    SenderDomainIs = @('contoso.com')
                    RecipientDomainIs = @('fabrikam.com')
                    SubjectContainsWords = @('invoice')
                    ExceptIfSubjectContainsWords = @('approved')
                }
            }
            """);

        var accepted = TransportRuleCommandBuilder.BuildTestTransportRuleScript(new TestTransportRuleRequest
        {
            Sender = "sender@contoso.com",
            Recipient = "user@fabrikam.com",
            Subject = "invoice needs review"
        });
        var excluded = TransportRuleCommandBuilder.BuildTestTransportRuleScript(new TestTransportRuleRequest
        {
            Sender = "sender@contoso.com",
            Recipient = "user@fabrikam.com",
            Subject = "approved invoice"
        });

        var acceptedResult = await fixture.Engine.ExecuteAsync(accepted);
        var excludedResult = await fixture.Engine.ExecuteAsync(excluded);

        Assert.True(acceptedResult.Success, acceptedResult.ErrorMessage);
        Assert.Equal("Allowed", Assert.Single(acceptedResult.Output).Properties["Name"]?.Value);
        Assert.True(excludedResult.Success, excludedResult.ErrorMessage);
        Assert.Empty(excludedResult.Output);
    }
}
