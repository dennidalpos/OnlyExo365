using ExchangeAdmin.Worker.PowerShell;

namespace ExchangeAdmin.Tests;

public sealed class PowerShellModuleBootstrapPolicyTests
{
    [Fact]
    public void Resolve_ReturnsAllowlistedModuleForExplicitRequestName()
    {
        var definition = PowerShellModuleBootstrapPolicy.Resolve("ExchangeOnlineManagement");

        Assert.NotNull(definition);
        Assert.Equal("ExchangeOnlineManagement", definition!.ModuleName);
        Assert.Equal("3.9.2", definition.RequiredVersion);
    }

    [Fact]
    public void Resolve_ReturnsAllowlistedModuleForAlias()
    {
        var definition = PowerShellModuleBootstrapPolicy.Resolve("Microsoft.Graph");

        Assert.NotNull(definition);
        Assert.Equal("Microsoft.Graph.Authentication", definition!.ModuleName);
        Assert.Equal("2.35.1", definition.RequiredVersion);
        Assert.Equal(
            [
                "Microsoft.Graph.Authentication",
                "Microsoft.Graph.Users",
                "Microsoft.Graph.Users.Actions",
                "Microsoft.Graph.Identity.DirectoryManagement"
            ],
            definition.GetRequiredModules());
    }

    [Fact]
    public void Resolve_ReturnsNullForUnknownModule()
    {
        Assert.Null(PowerShellModuleBootstrapPolicy.Resolve("Pester"));
    }

    [Fact]
    public void BuildManualInstructions_UsesEnglishBaselineText()
    {
        var definition = PowerShellModuleBootstrapPolicy.Resolve("ExchangeOnlineManagement");

        var instructions = definition!.BuildManualInstructions(PowerShellModuleBootstrapPolicy.RepositoryName);

        Assert.Contains("If available, run", instructions, StringComparison.Ordinal);
        Assert.Contains("Alternatively, download", instructions, StringComparison.Ordinal);
        Assert.DoesNotContain("Se disponibile", instructions, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("In alternativa", instructions, StringComparison.OrdinalIgnoreCase);
    }
}
