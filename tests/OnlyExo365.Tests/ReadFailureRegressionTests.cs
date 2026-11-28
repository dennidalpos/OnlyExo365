using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Tests;

public class ReadFailureRegressionTests
{
    [Theory]
    [InlineData("access denied")]
    [InlineData("transient service failure")]
    [InlineData("missing")]
    public async Task Connectors_FailedInboundDoesNotReturnSuccessfulEmptyOrOutboundOnly(string failure)
    {
        using var fixture = new LocalPowerShellFixture();
        await fixture.DefineCommandsAsync($$"""
            function Get-InboundConnector { [CmdletBinding()] param() throw '{{failure}}' }
            function Get-OutboundConnector { [CmdletBinding()] param() [pscustomobject]@{ Identity='out'; Name='out'; Enabled=$true } }
            """);
        if (failure == "missing") await fixture.DefineCommandsAsync(@"Remove-Item Function:\Get-InboundConnector");
        var error = await Assert.ThrowsAsync<InvalidOperationException>(() => new ExoMailFlowCommands(fixture.Engine).GetConnectorsAsync());
        Assert.Contains(failure == "missing" ? "Get-InboundConnector" : failure, error.Message);
    }

    [Fact]
    public async Task Connectors_EmptySuccessfulReadAndInformationalWarningRemainSuccessful()
    {
        using var fixture = new LocalPowerShellFixture();
        await fixture.DefineCommandsAsync("""
            function Get-InboundConnector { [CmdletBinding()] param() Write-Warning 'Informational only' }
            function Get-OutboundConnector { [CmdletBinding()] param() }
            """);
        Assert.Empty((await new ExoMailFlowCommands(fixture.Engine).GetConnectorsAsync()).Connectors);
    }

    [Fact]
    public async Task Connectors_UninitializedEngineDoesNotReturnSuccessfulEmptyRead()
    {
        using var engine = new PowerShellEngine();
        await Assert.ThrowsAsync<InvalidOperationException>(() => new ExoMailFlowCommands(engine).GetConnectorsAsync());
    }

    [Theory]
    [InlineData("tenant", "access denied")]
    [InlineData("available", "transient service failure")]
    [InlineData("user", "missing")]
    public async Task LicenseReads_PropagateErrorsInsteadOfEmptySuccess(string operation, string failure)
    {
        using var fixture = new LocalPowerShellFixture();
        fixture.Engine.GraphConnected = true;
        fixture.Engine.ConnectedGraphScopesValue = fixture.Engine.ExchangeConfiguration.NormalizeGraphScopes();
        await fixture.DefineCommandsAsync($$"""
            function Get-Module { [pscustomobject]@{ Version = [version]'2.41.2' } }
            function Import-Module { param($FullyQualifiedName) }
            function Get-MgSubscribedSku { [CmdletBinding()] param() throw '{{failure}}' }
            function Get-MgUserLicenseDetail { [CmdletBinding()] param($UserId) throw '{{failure}}' }
            """);
        if (failure == "missing") await fixture.DefineCommandsAsync(@"Remove-Item Function:\Get-MgUserLicenseDetail");
        var commands = new ExoMailboxLicenseCommands(fixture.Engine);
        var error = await Assert.ThrowsAsync<InvalidOperationException>(async () =>
        {
            if (operation == "tenant") await commands.GetTenantLicensesAsync();
            else if (operation == "available") await commands.GetAvailableLicensesAsync();
            else await commands.GetUserLicensesAsync("offline@example.test");
        });
        Assert.Contains(failure == "missing" ? "Get-MgUserLicenseDetail" : failure, error.Message);
    }
}
