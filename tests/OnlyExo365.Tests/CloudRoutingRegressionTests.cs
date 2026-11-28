using System.Management.Automation.Language;
using OnlyExo365.Contracts;
using OnlyExo365.Shell.Services;
using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Tests;

public class CloudRoutingRegressionTests
{
    [Theory]
    [InlineData("O365Default", "Global", "ps.compliance.protection.outlook.com", "login.microsoftonline.com")]
    [InlineData("O365USGovGCCHigh", "USGov", "ps.compliance.protection.office365.us", "login.microsoftonline.us")]
    [InlineData("O365USGovDoD", "USGovDoD", "l5.ps.compliance.protection.office365.us", "login.microsoftonline.us")]
    [InlineData("O365China", "China", "ps.compliance.protection.partner.outlook.cn", "login.chinacloudapi.cn")]
    public async Task AllModesAndEntryPoints_BindTheSelectedCloud(string cloud, string graph, string complianceHost, string authorityHost)
    {
        using var fixture = new LocalPowerShellFixture();
        await fixture.DefineCommandsAsync("""
            function Connect-MgGraph {
                [CmdletBinding()] param($Scopes,$TenantId,$ContextScope,[switch]$NoWelcome,[switch]$UseDeviceCode,[switch]$Identity,$ClientId,$CertificateThumbprint,$CertificateSubjectName,$Environment)
                $Environment
            }
            function Connect-IPPSSession {
                [CmdletBinding()] param($ConnectionUri,$AzureADAuthorizationEndpointUri,$Organization,$DelegatedOrganization,$UserPrincipalName,[switch]$Device,$AppId,$CertificateThumbprint,$CertificateSubjectName,[switch]$ManagedIdentity,$ManagedIdentityAccountId,[switch]$EnableSearchOnlySession,[bool]$ShowBanner)
                "$ConnectionUri|$AzureADAuthorizationEndpointUri"
            }
            """);
        var ast = Parser.ParseFile(TestPathHelper.GetRepositoryPath("build/run-tenant-validation.ps1"), out _, out var parseErrors);
        Assert.Empty(parseErrors);
        var helperPath = TestPathHelper.GetRepositoryPath("build/helpers/cloud-routing.ps1");
        await fixture.DefineCommandsAsync(File.ReadAllText(helperPath).Replace("$PSScriptRoot", "'" + Path.GetDirectoryName(helperPath)!.Replace("'", "''") + "'") + "\n" +
            string.Join("\n", ast.FindAll(node => node is FunctionDefinitionAst, false).Select(node => node.Extent.Text)));
        foreach (var mode in Enum.GetValues<ExchangeAuthenticationMode>())
        {
            var configuration = new ExchangeOnlineConfiguration
            {
                AuthenticationMode = mode, ExchangeEnvironmentName = cloud, ApplicationId = "offline-app",
                ExchangeOrganization = "example.onmicrosoft.com", GraphTenantId = "example.onmicrosoft.com", CertificateThumbprint = "FAKE"
            };
            var graphResult = await fixture.Engine.ExecuteAsync(GraphCommandBuilder.BuildConnectGraphCommand(configuration));
            Assert.True(graphResult.Success, graphResult.ErrorMessage);
            Assert.Equal(graph, Assert.Single(graphResult.Output).BaseObject.ToString());
            var supportsCompliance = mode is ExchangeAuthenticationMode.Interactive or ExchangeAuthenticationMode.AppCertificate;
            if (supportsCompliance)
            {
                var complianceResult = await fixture.Engine.ExecuteAsync(ComplianceCommandBuilder.BuildConnectComplianceCommand(configuration));
                Assert.True(complianceResult.Success, complianceResult.ErrorMessage);
                var routing = Assert.Single(complianceResult.Output).BaseObject.ToString()!;
                Assert.Contains(complianceHost, routing);
                Assert.Contains(authorityHost + "/organizations", routing);
            }
            else Assert.Throws<InvalidOperationException>(() => ComplianceCommandBuilder.BuildConnectComplianceCommand(configuration));
            var tenantResult = await fixture.Engine.ExecuteAsync($$"""
                $config = @{ ExchangeEnvironmentName='{{cloud}}'; GraphTenantId='example.onmicrosoft.com'; GraphScopes=@('Organization.Read.All');
                    ApplicationId='offline-app'; CertificateThumbprint='FAKE'; CertificateSubjectName=''; ExchangeOrganization='example.onmicrosoft.com';
                    DelegatedOrganization=''; ManagedIdentityAccountId=''; UserPrincipalNameHint='' }
                $inv = Get-GraphConnectCommand $config '{{mode}}'; $a = $inv.Arguments; & $inv.CommandName @a
                """);
            Assert.True(tenantResult.Success, tenantResult.ErrorMessage);
            Assert.Equal(graph, tenantResult.Output[0].BaseObject.ToString());
            var tenantCompliance = await fixture.Engine.ExecuteAsync($"$inv = Get-ComplianceConnectCommand $config '{mode}'; $a = $inv.Arguments; & $inv.CommandName @a");
            Assert.Equal(supportsCompliance, tenantCompliance.Success);
            if (supportsCompliance) Assert.Contains(complianceHost, Assert.Single(tenantCompliance.Output).BaseObject.ToString());
        }
        var bootstrap = InteractiveExchangeBootstrapService.BuildBootstrapScript(new() { ExchangeEnvironmentName = cloud, ApplicationId = "offline-app" });
        Assert.Contains($"-Environment '{graph}'", bootstrap);
        Assert.Contains(complianceHost, bootstrap);
        Assert.Contains(authorityHost, bootstrap);
        Parser.ParseInput(bootstrap, out _, out var errors);
        Assert.Empty(errors);
    }

    [Theory]
    [InlineData("O365GermanyCloud")]
    [InlineData("Unknown")]
    public void UnsupportedClouds_AreRejected(string cloud)
        => Assert.Throws<InvalidOperationException>(() => GraphCommandBuilder.BuildConnectGraphCommand(new() { ExchangeEnvironmentName = cloud }));

    [Fact]
    public void ChinaDelegatedGraph_RequiresCustomApplication()
        => Assert.Throws<InvalidOperationException>(() => GraphCommandBuilder.BuildConnectGraphCommand(new() { ExchangeEnvironmentName = "O365China" }));

    [Theory]
    [InlineData(ExchangeAuthenticationMode.DeviceCode)]
    [InlineData(ExchangeAuthenticationMode.ManagedIdentity)]
    [InlineData(ExchangeAuthenticationMode.AppCertificate)]
    public void SearchOnly_RejectsUnsupportedAuthentication(ExchangeAuthenticationMode mode)
        => Assert.Throws<InvalidOperationException>(() => ComplianceCommandBuilder.BuildConnectComplianceSearchOnlyCommand(new() { AuthenticationMode = mode }));
}
