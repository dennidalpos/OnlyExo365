using System.Management.Automation.Language;
using System.Text.Json;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Tests;

public class MailFlowWriteRegressionTests
{
    private static readonly string[] RuleArrays =
    [
        "From", "SentTo", "SenderDomainIs", "RecipientDomainIs", "SentToMemberOf",
        "SubjectContainsWords", "ExceptIfFrom", "ExceptIfSentTo", "ExceptIfSenderDomainIs",
        "ExceptIfRecipientDomainIs", "ExceptIfSubjectContainsWords", "RedirectMessageTo", "BlindCopyTo", "AddToRecipients"
    ];

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task RuleWrite_BindsAllExplicitValuesAndPreservesUnexposedFields(bool create, bool clear)
    {
        using var fixture = new LocalPowerShellFixture();
        await DefineRuleCommandsAsync(fixture);
        var request = new UpsertTransportRuleRequest
        {
            Identity = create ? null : "Rule'O'Brien", Name = "Finance'O'Brien",
            PrependSubject = clear ? "" : "[O'Brien]", StopRuleProcessing = !clear,
            DeleteMessage = !clear, Enabled = !clear, Mode = "Audit"
        };
        foreach (var field in RuleArrays)
            typeof(UpsertTransportRuleRequest).GetProperty(field)!.SetValue(request,
                clear ? new List<string>() : new List<string> { "O'Brien@example.test", "second@example.test" });

        // Exercise the IPC representation, including explicit empty arrays and false.
        request = JsonMessageSerializer.ExtractPayload<UpsertTransportRuleRequest>(JsonMessageSerializer.ToJsonElement(request))!;
        var script = TransportRuleCommandBuilder.BuildUpsertTransportRuleScript(request);
        Parser.ParseInput(script, out _, out var errors);
        Assert.Empty(errors);
        var result = await fixture.Engine.ExecuteAsync(script);
        Assert.True(result.Success, result.ErrorMessage);
        var calls = await ReadCallsAsync(fixture);
        var write = calls[0];
        Assert.Equal(create ? "New" : "Set", write.GetProperty("Command").GetString());
        var parameters = write.GetProperty("Parameters");
        Assert.Equal("Finance'O'Brien", parameters.GetProperty("Name").GetString());
        foreach (var field in RuleArrays)
        {
            if (clear && create) Assert.False(parameters.TryGetProperty(field, out _));
            else if (clear) Assert.Equal(JsonValueKind.Null, parameters.GetProperty(field).ValueKind);
            else Assert.Equal(new[] { "O'Brien@example.test", "second@example.test" },
                parameters.GetProperty(field).EnumerateArray().Select(v => v.GetString()));
        }
        if (clear && create) Assert.False(parameters.TryGetProperty("PrependSubject", out _));
        else Assert.Equal(clear ? null : "[O'Brien]", parameters.GetProperty("PrependSubject").GetString());
        Assert.Equal(!clear, parameters.GetProperty("StopRuleProcessing").GetBoolean());
        Assert.Equal(!clear, parameters.GetProperty("DeleteMessage").GetBoolean());
        Assert.Equal("Audit", parameters.GetProperty("Mode").GetString());
        Assert.False(parameters.TryGetProperty("ApplyHtmlDisclaimerText", out _));
        if (create)
        {
            Assert.Single(calls.EnumerateArray());
            Assert.Equal(!clear, parameters.GetProperty("Enabled").GetBoolean());
        }
        else
        {
            Assert.False(parameters.TryGetProperty("Enabled", out _));
            Assert.Equal(2, calls.GetArrayLength());
            Assert.Equal(clear ? "Disable" : "Enable", calls[1].GetProperty("Command").GetString());
            Assert.Equal("Rule'O'Brien", calls[1].GetProperty("Parameters").GetProperty("Identity").GetString());
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RuleWrite_OmittedValuesAreNotBoundAndCreateKeepsDefaults(bool create)
    {
        using var fixture = new LocalPowerShellFixture();
        await DefineRuleCommandsAsync(fixture);
        var request = JsonSerializer.Deserialize<UpsertTransportRuleRequest>(
            create ? "{\"name\":\"New\"}" : "{\"identity\":\"Existing\",\"name\":\"Renamed\"}")!;
        var result = await fixture.Engine.ExecuteAsync(TransportRuleCommandBuilder.BuildUpsertTransportRuleScript(request));
        Assert.True(result.Success, result.ErrorMessage);
        var calls = await ReadCallsAsync(fixture);
        Assert.Single(calls.EnumerateArray());
        var parameters = calls[0].GetProperty("Parameters");
        Assert.Equal(create ? 4 : 3, parameters.EnumerateObject().Count()); // includes ErrorAction
        if (create)
        {
            Assert.True(parameters.GetProperty("Enabled").GetBoolean());
            Assert.Equal("Enforce", parameters.GetProperty("Mode").GetString());
        }
    }

    [Fact]
    public async Task RuleWrite_FailureStopsBeforeStateChange()
    {
        using var fixture = new LocalPowerShellFixture();
        await DefineRuleCommandsAsync(fixture);
        await fixture.DefineCommandsAsync("function Set-TransportRule { throw 'Write rejected' }");
        var result = await fixture.Engine.ExecuteAsync(TransportRuleCommandBuilder.BuildUpsertTransportRuleScript(
            new UpsertTransportRuleRequest { Identity = "Existing", Name = "Name", Enabled = false }));
        Assert.False(result.Success);
        Assert.Contains("Write rejected", result.ErrorMessage);
        Assert.Empty((await ReadCallsAsync(fixture)).EnumerateArray());
    }

    [Theory]
    [InlineData("Inbound", false, "Partner")]
    [InlineData("Inbound", false, "OnPremises")]
    [InlineData("Outbound", false, "Partner")]
    [InlineData("Outbound", false, "OnPremises")]
    [InlineData("Inbound", true, "OnPremises")]
    [InlineData("Outbound", true, "OnPremises")]
    public async Task ConnectorWrite_BindsDirectionAndPreservesExistingCategory(string direction, bool create, string category)
    {
        using var fixture = new LocalPowerShellFixture();
        await DefineConnectorCommandsAsync(fixture, category);
        var commands = new ExoMailFlowCommands(fixture.Engine);
        await commands.UpsertConnectorAsync(new UpsertConnectorRequest
        {
            Identity = create ? null : "Connector'O'Brien", Type = direction, Name = "Name'O'Brien",
            Comment = "Comment'O'Brien", Enabled = false,
            SenderDomains = ["sender.example"], RecipientDomains = ["recipient.example"]
        });
        var calls = await ReadCallsAsync(fixture);
        var call = Assert.Single(calls.EnumerateArray());
        Assert.Equal((create ? "New-" : "Set-") + direction + "Connector", call.GetProperty("Command").GetString());
        var parameters = call.GetProperty("Parameters");
        Assert.Equal("Name'O'Brien", parameters.GetProperty("Name").GetString());
        Assert.Equal("Comment'O'Brien", parameters.GetProperty("Comment").GetString());
        Assert.False(parameters.GetProperty("Enabled").GetBoolean());
        var domainField = direction == "Inbound" ? "SenderDomains" : "RecipientDomains";
        Assert.Equal(direction == "Inbound" ? "sender.example" : "recipient.example", parameters.GetProperty(domainField)[0].GetString());
        Assert.False(parameters.TryGetProperty(direction == "Inbound" ? "RecipientDomains" : "SenderDomains", out _));
        Assert.Equal(category, call.GetProperty("Category").GetString());
        Assert.Equal(create, parameters.TryGetProperty("ConnectorType", out _));
        if (create) Assert.Equal("OnPremises", parameters.GetProperty("ConnectorType").GetString());
        else Assert.Equal("Connector'O'Brien", parameters.GetProperty("Identity").GetString());
    }

    [Theory]
    [InlineData("Inbound")]
    [InlineData("Outbound")]
    public async Task ConnectorDelete_BindsDirectionIdentityAndConfirmation(string direction)
    {
        using var fixture = new LocalPowerShellFixture();
        await DefineConnectorCommandsAsync(fixture, "Partner");
        await new ExoMailFlowCommands(fixture.Engine).RemoveConnectorAsync(new RemoveConnectorRequest
        { Identity = "Connector'O'Brien", Type = direction });
        var call = Assert.Single((await ReadCallsAsync(fixture)).EnumerateArray());
        Assert.Equal("Remove-" + direction + "Connector", call.GetProperty("Command").GetString());
        Assert.Equal("Connector'O'Brien", call.GetProperty("Parameters").GetProperty("Identity").GetString());
        Assert.False(call.GetProperty("Parameters").GetProperty("Confirm").GetProperty("IsPresent").GetBoolean());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task AcceptedDomainWrite_BindsCreateUpdateAndOptionalDefault(bool create, bool makeDefault)
    {
        using var fixture = new LocalPowerShellFixture();
        await fixture.DefineCommandsAsync("""
            $script:calls = [System.Collections.Generic.List[object]]::new()
            function Get-ConnectionInformation {
                [CmdletBinding()] param()
                [pscustomobject]@{ State = 'Connected'; UserPrincipalName = 'offline@example.test'; Organization = 'example.test' }
            }
            function New-AcceptedDomain {
                [CmdletBinding()] param([string]$Name, [string]$DomainName, [string]$DomainType)
                $script:calls.Add(@{ Command = 'New'; Parameters = $PSBoundParameters })
            }
            function Set-AcceptedDomain {
                [CmdletBinding()] param([string]$Identity, [string]$DomainType, [switch]$MakeDefault)
                $script:calls.Add(@{ Command = 'Set'; Parameters = $PSBoundParameters })
            }
            """);
        await new ExoMailFlowCommands(fixture.Engine).UpsertAcceptedDomainAsync(new UpsertAcceptedDomainRequest
        {
            Identity = create ? null : "Domain'O'Brien", Name = "Name'O'Brien", DomainName = "contoso.example",
            DomainType = "InternalRelay", MakeDefault = makeDefault
        });
        var calls = await ReadCallsAsync(fixture);
        Assert.Equal(makeDefault ? 2 : 1, calls.GetArrayLength());
        Assert.Equal(create ? "New" : "Set", calls[0].GetProperty("Command").GetString());
        var parameters = calls[0].GetProperty("Parameters");
        Assert.Equal("InternalRelay", parameters.GetProperty("DomainType").GetString());
        Assert.Equal(create ? "Name'O'Brien" : "Domain'O'Brien", parameters.GetProperty(create ? "Name" : "Identity").GetString());
        if (create) Assert.Equal("contoso.example", parameters.GetProperty("DomainName").GetString());
        if (makeDefault)
        {
            Assert.Equal("Set", calls[1].GetProperty("Command").GetString());
            Assert.Equal("contoso.example", calls[1].GetProperty("Parameters").GetProperty("Identity").GetString());
            Assert.True(calls[1].GetProperty("Parameters").GetProperty("MakeDefault").GetProperty("IsPresent").GetBoolean());
        }
    }

    private static Task DefineRuleCommandsAsync(LocalPowerShellFixture fixture) => fixture.DefineCommandsAsync("""
        $script:calls = [System.Collections.Generic.List[object]]::new()
        function Set-TransportRule {
            [CmdletBinding()]
            param([string]$Identity, [string]$Name, [string]$Mode,
                [string[]]$From, [string[]]$SentTo, [string[]]$SenderDomainIs, [string[]]$RecipientDomainIs,
                [string[]]$SentToMemberOf, [string[]]$SubjectContainsWords, [string[]]$ExceptIfFrom,
                [string[]]$ExceptIfSentTo, [string[]]$ExceptIfSenderDomainIs, [string[]]$ExceptIfRecipientDomainIs,
                [string[]]$ExceptIfSubjectContainsWords, [string[]]$RedirectMessageTo, [string[]]$BlindCopyTo,
                [string[]]$AddToRecipients, [AllowNull()][object]$PrependSubject,
                [bool]$StopRuleProcessing, [bool]$DeleteMessage)
            $script:calls.Add(@{ Command = 'Set'; Parameters = $PSBoundParameters })
        }
        function New-TransportRule {
            [CmdletBinding()]
            param([string]$Name, [string]$Mode, [bool]$Enabled,
                [string[]]$From, [string[]]$SentTo, [string[]]$SenderDomainIs, [string[]]$RecipientDomainIs,
                [string[]]$SentToMemberOf, [string[]]$SubjectContainsWords, [string[]]$ExceptIfFrom,
                [string[]]$ExceptIfSentTo, [string[]]$ExceptIfSenderDomainIs, [string[]]$ExceptIfRecipientDomainIs,
                [string[]]$ExceptIfSubjectContainsWords, [string[]]$RedirectMessageTo, [string[]]$BlindCopyTo,
                [string[]]$AddToRecipients, [string]$PrependSubject, [bool]$StopRuleProcessing, [bool]$DeleteMessage)
            $script:calls.Add(@{ Command = 'New'; Parameters = $PSBoundParameters })
        }
        function Enable-TransportRule {
            [CmdletBinding(SupportsShouldProcess)] param([string]$Identity)
            $script:calls.Add(@{ Command = 'Enable'; Parameters = $PSBoundParameters })
        }
        function Disable-TransportRule {
            [CmdletBinding(SupportsShouldProcess)] param([string]$Identity)
            $script:calls.Add(@{ Command = 'Disable'; Parameters = $PSBoundParameters })
        }
        """);

    private static Task DefineConnectorCommandsAsync(LocalPowerShellFixture fixture, string category) => fixture.DefineCommandsAsync($$"""
        $script:calls = [System.Collections.Generic.List[object]]::new()
        $script:category = '{{category}}'
        function Record-Connector($command, $parameters) {
            if ($parameters.ContainsKey('ConnectorType')) { $script:category = $parameters['ConnectorType'] }
            $script:calls.Add(@{ Command = $command; Parameters = $parameters; Category = $script:category })
        }
        function Set-InboundConnector {
            [CmdletBinding()] param([string]$Identity, [string]$Name, [bool]$Enabled, [string]$Comment, [string[]]$SenderDomains, [string]$ConnectorType)
            Record-Connector 'Set-InboundConnector' $PSBoundParameters
        }
        function New-InboundConnector {
            [CmdletBinding()] param([string]$Name, [bool]$Enabled, [string]$Comment, [string[]]$SenderDomains, [string]$ConnectorType)
            Record-Connector 'New-InboundConnector' $PSBoundParameters
        }
        function Set-OutboundConnector {
            [CmdletBinding()] param([string]$Identity, [string]$Name, [bool]$Enabled, [string]$Comment, [string[]]$RecipientDomains, [string]$ConnectorType)
            Record-Connector 'Set-OutboundConnector' $PSBoundParameters
        }
        function New-OutboundConnector {
            [CmdletBinding()] param([string]$Name, [bool]$Enabled, [string]$Comment, [string[]]$RecipientDomains, [string]$ConnectorType)
            Record-Connector 'New-OutboundConnector' $PSBoundParameters
        }
        function Remove-InboundConnector {
            [CmdletBinding(SupportsShouldProcess)] param([string]$Identity)
            Record-Connector 'Remove-InboundConnector' $PSBoundParameters
        }
        function Remove-OutboundConnector {
            [CmdletBinding(SupportsShouldProcess)] param([string]$Identity)
            Record-Connector 'Remove-OutboundConnector' $PSBoundParameters
        }
        """);

    private static async Task<JsonElement> ReadCallsAsync(LocalPowerShellFixture fixture)
    {
        var result = await fixture.Engine.ExecuteAsync("ConvertTo-Json -InputObject @($script:calls.ToArray()) -Depth 8 -Compress");
        Assert.True(result.Success, result.ErrorMessage);
        using var document = JsonDocument.Parse(Assert.Single(result.Output).BaseObject.ToString()!);
        return document.RootElement.Clone();
    }
}
