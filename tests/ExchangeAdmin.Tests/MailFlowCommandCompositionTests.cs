using System.Reflection;
using ExchangeAdmin.Worker.PowerShell;

namespace ExchangeAdmin.Tests;

public class MailFlowCommandCompositionTests
{
    [Fact]
    public void ExoMailFlowCommands_ComposesDedicatedModulesPerMailFlowArea()
    {
        using var engine = new PowerShellEngine();
        var facade = new ExoMailFlowCommands(engine);

        AssertFieldType(facade, "_messageTraceCommands", "ExoMessageTraceCommands");
        AssertFieldType(facade, "_transportRuleCommands", "ExoTransportRuleCommands");
        AssertFieldType(facade, "_connectorCommands", "ExoConnectorCommands");
        AssertFieldType(facade, "_acceptedDomainCommands", "ExoAcceptedDomainCommands");
        AssertFieldType(facade, "_remoteDomainCommands", "ExoRemoteDomainCommands");
        AssertFieldType(facade, "_organizationRelationshipCommands", "ExoOrganizationRelationshipCommands");
        AssertFieldType(facade, "_addressListCommands", "ExoAddressListCommands");
        AssertFieldType(facade, "_addressBookPolicyCommands", "ExoAddressBookPolicyCommands");
        AssertFieldType(facade, "_offlineAddressBookCommands", "ExoOfflineAddressBookCommands");
        AssertFieldType(facade, "_sharingPolicyCommands", "ExoSharingPolicyCommands");
    }

    [Fact]
    public void BuildGetAddressListsScript_EmitsStructuredWarningWithValidConcatenation()
    {
        var script = ExoAddressListCommands.BuildGetAddressListsScript();

        Assert.Contains("Write-Warning ('__EA_WARN__' + $warningPayload)", script, StringComparison.Ordinal);
        Assert.DoesNotContain("Write-Warning '__EA_WARN__' + $warningPayload", script, StringComparison.Ordinal);
        Assert.Contains("Get-Command -Name 'Get-AddressList'", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildGetOfflineAddressBooksScript_EmitsStructuredWarningWithValidConcatenation()
    {
        var script = ExoOfflineAddressBookCommands.BuildGetOfflineAddressBooksScript();

        Assert.Contains("Write-Warning ('__EA_WARN__' + $warningPayload)", script, StringComparison.Ordinal);
        Assert.DoesNotContain("Write-Warning '__EA_WARN__' + $warningPayload", script, StringComparison.Ordinal);
        Assert.Contains("Get-Command -Name 'Get-OfflineAddressBook'", script, StringComparison.Ordinal);
    }

    private static void AssertFieldType(object instance, string fieldName, string expectedTypeName)
    {
        var field = instance.GetType().GetField(fieldName, BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(field);

        var value = field!.GetValue(instance);
        Assert.NotNull(value);
        Assert.Equal(expectedTypeName, value!.GetType().Name);
    }
}
