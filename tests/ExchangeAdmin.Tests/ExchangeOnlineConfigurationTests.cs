using System.Collections.Specialized;
using ExchangeAdmin.Contracts;

namespace ExchangeAdmin.Tests;

public class ExchangeOnlineConfigurationTests
{
    [Fact]
    public void CreateDefault_UsesLeastPrivilegeGraphDefaultsAndConnectsGraphDuringInitialBootstrap()
    {
        var configuration = ExchangeOnlineConfiguration.CreateDefault();

        Assert.Equal(
            new[]
            {
                "Organization.Read.All",
                "Directory.Read.All",
                "RoleManagement.Read.Directory",
                "User.Read.All"
            },
            configuration.GraphScopes);
        Assert.Empty(configuration.NormalizeGraphLicenseWriteScopes());
        Assert.True(configuration.EnableGraphAfterExchangeConnect);
    }

    [Fact]
    public void NormalizeGraphScopes_DeduplicatesAndTrimsValues()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            GraphScopes = new List<string> { " User.Read.All ", "user.read.all", "", "Mail.Read" }
        };

        var scopes = configuration.NormalizeGraphScopes();

        Assert.Equal(new[] { "User.Read.All", "Mail.Read" }, scopes);
    }

    [Fact]
    public void ApplyEnvironmentVariables_WritesNormalizedValues()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            ExchangeEnvironmentName = "O365USGovDoD",
            AuthenticationMode = ExchangeAuthenticationMode.Interactive,
            ExchangeOrganization = "contoso.onmicrosoft.com",
            DelegatedOrganization = "delegated.onmicrosoft.com",
            UserPrincipalNameHint = "admin@contoso.com",
            GraphTenantId = "tenant-id",
            GraphScopes = new List<string> { " User.Read.All ", "User.Read.All", "Mail.Read" },
            GraphLicenseWriteScopes = new List<string> { " LicenseAssignment.ReadWrite.All ", "licenseassignment.readwrite.all" },
            DefaultUsageLocation = " it ",
            EnableGraphAfterExchangeConnect = true
        };

        var environmentVariables = new StringDictionary();

        configuration.ApplyEnvironmentVariables(environmentVariables);

        Assert.Equal("O365USGovDoD", environmentVariables[ExchangeConfigurationEnvironmentVariables.ExchangeEnvironmentName]);
        Assert.Equal("Interactive", environmentVariables[ExchangeConfigurationEnvironmentVariables.AuthenticationMode]);
        Assert.Equal("User.Read.All;Mail.Read", environmentVariables[ExchangeConfigurationEnvironmentVariables.GraphScopes]);
        Assert.Equal("LicenseAssignment.ReadWrite.All", environmentVariables[ExchangeConfigurationEnvironmentVariables.GraphLicenseWriteScopes]);
        Assert.Equal("IT", environmentVariables[ExchangeConfigurationEnvironmentVariables.DefaultUsageLocation]);
        Assert.Equal("1", environmentVariables[ExchangeConfigurationEnvironmentVariables.EnableGraphAfterExchangeConnect]);
    }

    [Fact]
    public void FromEnvironmentVariables_ReadsOverridesAndFallsBackToDefaults()
    {
        var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase)
        {
            [ExchangeConfigurationEnvironmentVariables.ExchangeEnvironmentName] = " O365GermanyCloud ",
            [ExchangeConfigurationEnvironmentVariables.AuthenticationMode] = "DeviceCode",
            [ExchangeConfigurationEnvironmentVariables.ExchangeOrganization] = " tenant.onmicrosoft.com ",
            [ExchangeConfigurationEnvironmentVariables.ApplicationId] = " app-id ",
            [ExchangeConfigurationEnvironmentVariables.GraphScopes] = "User.Read; Mail.Read ",
            [ExchangeConfigurationEnvironmentVariables.GraphLicenseWriteScopes] = "LicenseAssignment.ReadWrite.All",
            [ExchangeConfigurationEnvironmentVariables.DefaultUsageLocation] = " de ",
            [ExchangeConfigurationEnvironmentVariables.EnableGraphAfterExchangeConnect] = "true"
        };

        var configuration = ExchangeOnlineConfiguration.FromEnvironmentVariables(key =>
            values.TryGetValue(key, out var value) ? value : null);

        Assert.Equal("O365GermanyCloud", configuration.ExchangeEnvironmentName);
        Assert.Equal(ExchangeAuthenticationMode.DeviceCode, configuration.AuthenticationMode);
        Assert.Equal("tenant.onmicrosoft.com", configuration.ExchangeOrganization);
        Assert.Equal("app-id", configuration.ApplicationId);
        Assert.Equal(new[] { "User.Read", "Mail.Read" }, configuration.GraphScopes);
        Assert.Equal(new[] { "LicenseAssignment.ReadWrite.All" }, configuration.NormalizeGraphLicenseWriteScopes());
        Assert.Equal("DE", configuration.DefaultUsageLocation);
        Assert.True(configuration.EnableGraphAfterExchangeConnect);
        Assert.Null(configuration.DelegatedOrganization);
    }

    [Fact]
    public void FromEnvironmentVariables_ThrowsWhenAuthenticationModeIsInvalid()
    {
        var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase)
        {
            [ExchangeConfigurationEnvironmentVariables.AuthenticationMode] = "NotARealMode"
        };

        var exception = Assert.Throws<InvalidOperationException>(() =>
            ExchangeOnlineConfiguration.FromEnvironmentVariables(key =>
                values.TryGetValue(key, out var value) ? value : null));

        Assert.Contains("AuthenticationMode must be one of:", exception.Message);
    }

    [Fact]
    public void GetGraphScopesForLicenseWrite_AppendsConfiguredWriteScopesToReadOnlyDefaults()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            GraphScopes = new List<string> { "Organization.Read.All", "User.Read.All" },
            GraphLicenseWriteScopes = new List<string> { "LicenseAssignment.ReadWrite.All", "user.read.all" }
        };

        var scopes = configuration.GetGraphScopesForLicenseWrite();

        Assert.Equal(
            new[]
            {
                "Organization.Read.All",
                "User.Read.All",
                "LicenseAssignment.ReadWrite.All"
            },
            scopes);
    }

    [Fact]
    public void Validate_AppCertificateRequiresApplicationAndCertificateContext()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.AppCertificate,
            EnableGraphAfterExchangeConnect = true
        };

        var errors = configuration.Validate();

        Assert.Contains("ApplicationId is required for AppCertificate authentication.", errors);
        Assert.Contains("ExchangeOrganization is required for AppCertificate authentication.", errors);
        Assert.Contains("CertificateThumbprint or CertificateSubjectName is required for AppCertificate authentication.", errors);
        Assert.Contains("GraphTenantId is required when Graph is enabled with AppCertificate authentication.", errors);
    }

    [Fact]
    public void Validate_ManagedIdentityRequiresOrganizationAndRejectsCertificateFields()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.ManagedIdentity,
            CertificateThumbprint = "thumb"
        };

        var errors = configuration.Validate();

        Assert.Contains("ExchangeOrganization is required for ManagedIdentity authentication.", errors);
        Assert.Contains("Certificate settings are not applicable to ManagedIdentity authentication.", errors);
    }

    [Fact]
    public void Validate_RejectsInvalidOrganizationAndGraphTenantFormats()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            ExchangeOrganization = "not a tenant",
            GraphTenantId = "tenant id with spaces"
        };

        var errors = configuration.Validate();

        Assert.Contains("ExchangeOrganization must be a tenant domain like 'contoso.onmicrosoft.com' or a tenant GUID.", errors);
        Assert.Contains("GraphTenantId must be a tenant domain like 'contoso.onmicrosoft.com' or a tenant GUID.", errors);
    }

    [Fact]
    public void Validate_RejectsInvalidDefaultUsageLocation()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            DefaultUsageLocation = "ITA"
        };

        var errors = configuration.Validate();

        Assert.Contains("DefaultUsageLocation must be a valid two-letter country/region code.", errors);
    }
}
