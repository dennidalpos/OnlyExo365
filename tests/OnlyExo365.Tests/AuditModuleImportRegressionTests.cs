using OnlyExo365.Contracts;
using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Tests;

public sealed class AuditModuleImportRegressionTests
{
    [Theory]
    [InlineData("Microsoft.Graph", 4)]
    [InlineData("ExchangeOnlineManagement", 1)]
    public async Task Imports_SelectApprovedVersionEvenWhenNewerIsInstalled(string module, int expectedImports)
    {
        using var fixture = new LocalPowerShellFixture();
        var version = PowerShellModuleBootstrapPolicy.Resolve(module)!.RequiredVersion;
        await fixture.DefineCommandsAsync($$"""
            $global:approvedVersion = '{{version}}'
            $global:imports = 0
            function Get-Module {
                [pscustomobject]@{ Version = [version]'9.9.9' }
                [pscustomobject]@{ Version = [version]$global:approvedVersion }
            }
            function Import-Module {
                param($FullyQualifiedName)
                if ($null -eq $FullyQualifiedName -or $FullyQualifiedName.RequiredVersion -ne $global:approvedVersion) { throw 'Approved version not selected' }
                $global:imports++
            }
            """);
        var result = await fixture.Engine.ExecuteAsync(PowerShellModuleBootstrapPolicy.Resolve(module)!.BuildImportScript());
        Assert.True(result.Success, result.ErrorMessage);
        var imports = await fixture.Engine.ExecuteAsync("$global:imports");
        Assert.Equal(expectedImports, Convert.ToInt32(imports.Output.Single().BaseObject));
    }
    [Fact]
    public async Task MissingApprovedModule_FailsBeforeImportingNewerVersion()
    {
        using var fixture = new LocalPowerShellFixture();
        await fixture.DefineCommandsAsync("""
            function Get-Module { [pscustomobject]@{ Version = [version]'9.9.9' } }
            function Import-Module { throw 'Unapproved import attempted' }
            """);
        var result = await fixture.Engine.ExecuteAsync(ExoMailboxLicenseCommands.BuildImportRequiredGraphModulesScript());
        Assert.False(result.Success);
        Assert.Contains("is not installed. Install it from Tools", result.ErrorMessage, StringComparison.Ordinal);
    }
}
