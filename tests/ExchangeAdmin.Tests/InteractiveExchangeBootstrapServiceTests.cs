using ExchangeAdmin.Contracts;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Tests;

public sealed class InteractiveExchangeBootstrapServiceTests
{
    [Fact]
    public void BuildBootstrapScript_UsesInteractiveExchangeGraphAndComplianceCommands()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.Interactive,
            ExchangeEnvironmentName = "O365Default",
            ExchangeOrganization = "contoso.onmicrosoft.com",
            DelegatedOrganization = "delegated.contoso.com",
            UserPrincipalNameHint = "admin@contoso.com",
            GraphTenantId = "contoso.onmicrosoft.com",
            GraphScopes = ["User.Read.All", "Directory.Read.All"]
        };

        var script = InteractiveExchangeBootstrapService.BuildBootstrapScript(configuration);

        Assert.Contains("Connect-ExchangeOnline -ShowBanner:$false", script, StringComparison.Ordinal);
        Assert.Contains("-ExchangeEnvironmentName 'O365Default'", script, StringComparison.Ordinal);
        Assert.Contains("-Organization 'contoso.onmicrosoft.com'", script, StringComparison.Ordinal);
        Assert.Contains("-DelegatedOrganization 'delegated.contoso.com'", script, StringComparison.Ordinal);
        Assert.Contains("-UserPrincipalName 'admin@contoso.com'", script, StringComparison.Ordinal);
        Assert.Contains("Connect-MgGraph", script, StringComparison.Ordinal);
        Assert.Contains("-TenantId 'contoso.onmicrosoft.com'", script, StringComparison.Ordinal);
        Assert.Contains("Connect-IPPSSession", script, StringComparison.Ordinal);
        Assert.Contains("Save-Status -Payload $status", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildBootstrapScript_SkipsUnsupportedExchangeEnvironmentName()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.Interactive,
            ExchangeEnvironmentName = "UnsupportedCloud"
        };

        var script = InteractiveExchangeBootstrapService.BuildBootstrapScript(configuration);

        Assert.DoesNotContain("-ExchangeEnvironmentName 'UnsupportedCloud'", script, StringComparison.Ordinal);
    }
}
