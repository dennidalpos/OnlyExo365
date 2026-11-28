using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Tests;

public class RuleMatcherRegressionTests
{
    [Theory]
    [InlineData("Disabled", "Enforce", "", "NoMatch")]
    [InlineData("Enabled", "Audit", "", "Indeterminate")]
    [InlineData("Enabled", "Enforce", "SentToMemberOf=@('Group')", "Indeterminate")]
    [InlineData("Enabled", "Enforce", "Conditions=@([pscustomobject]@{Name='SentToScope'})", "Indeterminate")]
    [InlineData("Enabled", "Enforce", "ExceptIfSenderDomainIs=@('example.test')", "NoMatch")]
    [InlineData("Enabled", "Enforce", "SenderDomainIs=@('other.test')", "NoMatch")]
    [InlineData("Enabled", "Enforce", "From=@('Display Name')", "Indeterminate")]
    [InlineData("Enabled", "Enforce", "SubjectContainsWords=@('inv*')", "NoMatch")]
    [InlineData("Enabled", "Enforce", "", "Match")]
    public async Task Preview_ReportsUncertaintyStateAndExceptions(string state, string mode, string extra, string expected)
    {
        using var fixture = new LocalPowerShellFixture();
        await fixture.DefineCommandsAsync($$"""
            function Get-TransportRule {
                [CmdletBinding()] param()
                $rule = [pscustomobject]@{Name='Rule'; State='{{state}}'; Mode='{{mode}}'; Conditions=@(); Exceptions=@(); SenderDomainIs=@('example.test')}
                $extra = @{ {{extra}} }
                foreach ($key in $extra.Keys) { $rule | Add-Member -NotePropertyName $key -NotePropertyValue $extra[$key] -Force }
                $rule
            }
            """);
        var result = await new ExoMailFlowCommands(fixture.Engine).TestTransportRuleAsync(new()
        { Sender = "sender@example.test", Recipient = "recipient@example.test", Subject = "invoice" });
        var evaluation = Assert.Single(result.Evaluations);
        Assert.Equal(expected, evaluation.Status.ToString());
        if (expected == "Indeterminate") Assert.NotEmpty(evaluation.Reasons);
    }

    [Fact]
    public async Task Preview_MissingMetadataIsIndeterminate()
    {
        using var fixture = new LocalPowerShellFixture();
        await fixture.DefineCommandsAsync("function Get-TransportRule { [pscustomobject]@{Name='Unknown';State='Enabled';Mode='Enforce'} }");
        var result = await new ExoMailFlowCommands(fixture.Engine).TestTransportRuleAsync(new());
        Assert.Equal(TransportRuleMatchStatus.Indeterminate, Assert.Single(result.Evaluations).Status);
    }

    [Fact]
    public async Task Preview_PropagatesReadFailure()
    {
        using var fixture = new LocalPowerShellFixture();
        await fixture.DefineCommandsAsync("function Get-TransportRule { throw 'access denied' }");
        var error = await Assert.ThrowsAsync<InvalidOperationException>(() => new ExoMailFlowCommands(fixture.Engine).TestTransportRuleAsync(new()));
        Assert.Contains("access denied", error.Message);
    }
}
