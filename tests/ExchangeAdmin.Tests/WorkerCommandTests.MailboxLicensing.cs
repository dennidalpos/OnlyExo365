using ExchangeAdmin.Contracts;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Worker;
using ExchangeAdmin.Worker.PowerShell;

namespace ExchangeAdmin.Tests;

public partial class WorkerCommandTests
{
    [Fact]
    public void BuildConnectGraphCommand_BuildsDeviceCodeFlow()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.DeviceCode,
            GraphTenantId = "contoso.onmicrosoft.com",
            GraphScopes = new List<string> { "User.Read.All" }
        };

        var command = GraphCommandBuilder.BuildConnectGraphCommand(
            configuration,
            delegatedScopes: configuration.GetGraphScopesForLicenseWrite());

        Assert.Contains("-UseDeviceCode", command, StringComparison.Ordinal);
        Assert.Contains("-TenantId 'contoso.onmicrosoft.com'", command, StringComparison.Ordinal);
        Assert.Contains("-Scopes @('User.Read.All')", command, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildConnectGraphCommand_UsesExplicitDelegatedScopeSetForWriteEscalation()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.Interactive,
            GraphTenantId = "contoso.onmicrosoft.com",
            GraphScopes = new List<string> { "User.Read.All", "Directory.Read.All" },
            GraphLicenseWriteScopes = new List<string> { "LicenseAssignment.ReadWrite.All" }
        };

        var command = GraphCommandBuilder.BuildConnectGraphCommand(
            configuration,
            delegatedScopes: configuration.GetGraphScopesForLicenseWrite());

        Assert.Contains("-TenantId 'contoso.onmicrosoft.com'", command, StringComparison.Ordinal);
        Assert.Contains("-Scopes @('User.Read.All', 'Directory.Read.All', 'LicenseAssignment.ReadWrite.All')", command, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildSetUserLicenseScript_ValidatesUsageLocationBeforeAddingLicenses()
    {
        var script = ExoMailboxLicenseCommands.BuildSetUserLicenseScript(new SetUserLicenseRequest
        {
            UserPrincipalName = "mario.rossi@contoso.com",
            AddLicenseSkuIds = ["sku-1"]
        });

        Assert.Contains("Microsoft.Graph.Users.Actions", script, StringComparison.Ordinal);
        Assert.Contains("Get-MgUser -UserId $userId -Property Id,DisplayName,UserPrincipalName,UsageLocation -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("if ([string]::IsNullOrWhiteSpace($usageLocation))", script, StringComparison.Ordinal);
        Assert.Contains("if ($usageLocation -notmatch '^[A-Za-z]{2}$')", script, StringComparison.Ordinal);
        Assert.Contains("Set-MgUserLicense -UserId $userId -AddLicenses $addLicenses -RemoveLicenses $removeLicenses -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("invalid usage location", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildSetUserLicenseScript_SkipsUsageLocationValidationWhenOnlyRemovingLicenses()
    {
        var script = ExoMailboxLicenseCommands.BuildSetUserLicenseScript(new SetUserLicenseRequest
        {
            UserPrincipalName = "mario.rossi@contoso.com",
            RemoveLicenseSkuIds = ["sku-1"]
        });

        Assert.Contains("$shouldValidateUsageLocation = $false", script, StringComparison.Ordinal);
        Assert.Contains("if ($shouldValidateUsageLocation) {", script, StringComparison.Ordinal);
        Assert.Contains("Set-MgUserLicense -UserId $userId -AddLicenses $addLicenses -RemoveLicenses $removeLicenses -ErrorAction Stop", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildGetUsageLocationSuggestionScript_UsesTenantThenConfigurationFallback()
    {
        var script = ExoMailboxLicenseCommands.BuildGetUsageLocationSuggestionScript(
            new GetUsageLocationSuggestionRequest
            {
                UserPrincipalName = "mario.rossi@contoso.com"
            },
            "it");

        Assert.Contains("Microsoft.Graph.Identity.DirectoryManagement", script, StringComparison.Ordinal);
        Assert.Contains("Microsoft.Graph.Users", script, StringComparison.Ordinal);
        Assert.Contains("Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1", script, StringComparison.Ordinal);
        Assert.Contains("$suggestionSource = 'Tenant'", script, StringComparison.Ordinal);
        Assert.Contains("$suggestionSource = 'Configuration'", script, StringComparison.Ordinal);
        Assert.Contains("$fallbackUsageLocation = 'IT'", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildImportRequiredGraphModulesScript_RequiresLicensingCommandModules()
    {
        var script = ExoMailboxLicenseCommands.BuildImportRequiredGraphModulesScript();

        Assert.Contains("Microsoft.Graph.Authentication", script, StringComparison.Ordinal);
        Assert.Contains("Microsoft.Graph.Users", script, StringComparison.Ordinal);
        Assert.Contains("Microsoft.Graph.Users.Actions", script, StringComparison.Ordinal);
        Assert.Contains("Microsoft.Graph.Identity.DirectoryManagement", script, StringComparison.Ordinal);
        Assert.Contains("Install the approved Graph bundle from Tools and retry.", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildGetAdminRoleMembersScript_EmitsStructuredWarningsInsteadOfSwallowingRoleFailures()
    {
        var script = ExoMailboxLicenseCommands.BuildGetAdminRoleMembersScript();

        Assert.Contains("function Write-AdminRoleWarning", script, StringComparison.Ordinal);
        Assert.Contains("Write-Warning \"__EA_WARN__", script, StringComparison.Ordinal);
        Assert.Contains("AdminRoleMembersLoadFailed", script, StringComparison.Ordinal);
        Assert.Contains("AdminRoleUserLoadFailed", script, StringComparison.Ordinal);
        Assert.Contains("AdminRoleDirectoryQueryFailed", script, StringComparison.Ordinal);
        Assert.DoesNotContain("catch {}", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildSetUserUsageLocationScript_UpdatesGraphUserWithNormalizedCountryCode()
    {
        var script = ExoMailboxLicenseCommands.BuildSetUserUsageLocationScript(new SetUserUsageLocationRequest
        {
            UserPrincipalName = "mario.rossi@contoso.com",
            UsageLocation = " de "
        });

        Assert.Contains("Update-MgUser -UserId $userId -UsageLocation $usageLocation -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("$usageLocation = 'DE'", script, StringComparison.Ordinal);
    }

    [Fact]
    public void LicenseSkuNameResolver_UsesEmbeddedMicrosoftCatalogForDisplayNames()
    {
        var displayName = LicenseSkuNameResolver.Resolve("ENTERPRISEPACK", "6fd2c87f-b296-42f0-b197-1e91e994b900");

        Assert.Equal("Office 365 E3", displayName);
    }

    [Fact]
    public void LicenseSkuNameResolver_DetectsExchangeOnlineServicePlansFromCatalog()
    {
        Assert.True(LicenseSkuNameResolver.HasExchangeOnlineServicePlan("ENTERPRISEPACK", "6fd2c87f-b296-42f0-b197-1e91e994b900"));
        Assert.False(LicenseSkuNameResolver.HasExchangeOnlineServicePlan("POWER_BI_PRO"));
    }

    [Fact]
    public void LicenseSkuNameResolver_ReturnsCatalogServicePlans()
    {
        var servicePlans = LicenseSkuNameResolver.GetServicePlans("ENTERPRISEPACK", "6fd2c87f-b296-42f0-b197-1e91e994b900");

        Assert.Contains(servicePlans, plan => plan.ServicePlanName == "EXCHANGE_S_ENTERPRISE");
        Assert.Contains(servicePlans, plan => plan.FriendlyName == "Exchange Online (Plan 2)");
    }
}
