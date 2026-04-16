using ExchangeAdmin.Contracts;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Tests;

public sealed class ExchangeConfigurationLoaderTests : IDisposable
{
    private readonly string _tempDirectory = Path.Combine(Path.GetTempPath(), "ExchangeAdmin.Tests", Guid.NewGuid().ToString("N"));
    private readonly string _sharedConfigurationDirectory;

    public ExchangeConfigurationLoaderTests()
    {
        Directory.CreateDirectory(_tempDirectory);
        _sharedConfigurationDirectory = Path.Combine(_tempDirectory, "ProgramData");
        Directory.CreateDirectory(_sharedConfigurationDirectory);
    }

    [Fact]
    public void Load_ThrowsWhenAppSettingsJsonIsMalformed()
    {
        File.WriteAllText(Path.Combine(_tempDirectory, "appsettings.json"), "{ invalid json");

        var exception = Assert.Throws<ExchangeConfigurationLoader.ExchangeConfigurationLoadException>(
            () => ExchangeConfigurationLoader.Load(_tempDirectory, _sharedConfigurationDirectory, _ => null));

        Assert.Contains("Invalid JSON", exception.Message);
        Assert.Contains("appsettings.json", exception.FilePath);
    }

    [Fact]
    public void Load_ThrowsWhenAppSettingsJsonProducesInvalidConfiguration()
    {
        File.WriteAllText(
            Path.Combine(_tempDirectory, "appsettings.json"),
            """
            {
              "exchangeOnline": {
                "authenticationMode": "AppCertificate",
                "exchangeEnvironmentName": "O365Default",
                "enableGraphAfterExchangeConnect": true
              }
            }
            """);

        var exception = Assert.Throws<ExchangeConfigurationLoader.ExchangeConfigurationLoadException>(
            () => ExchangeConfigurationLoader.Load(_tempDirectory, _sharedConfigurationDirectory, _ => null));

        Assert.Contains("Invalid Exchange configuration", exception.Message);
        Assert.Contains("ApplicationId is required", exception.Message);
    }

    [Fact]
    public void Load_ReturnsConfigurationWhenAppSettingsJsonIsValid()
    {
        File.WriteAllText(
            Path.Combine(_tempDirectory, "appsettings.json"),
            """
            {
              "exchangeOnline": {
                "authenticationMode": "Interactive",
                "exchangeEnvironmentName": "O365Default",
                "graphScopes": [ "Organization.Read.All", "Directory.Read.All" ],
                "graphLicenseWriteScopes": [ "LicenseAssignment.ReadWrite.All" ],
                "enableGraphAfterExchangeConnect": true
              }
            }
            """);

        var configuration = ExchangeConfigurationLoader.Load(_tempDirectory, _sharedConfigurationDirectory, _ => null);

        Assert.Equal(ExchangeAuthenticationMode.Interactive, configuration.AuthenticationMode);
        Assert.True(configuration.EnableGraphAfterExchangeConnect);
        Assert.Equal(2, configuration.GraphScopes.Count);
        Assert.Equal(new[] { "LicenseAssignment.ReadWrite.All" }, configuration.NormalizeGraphLicenseWriteScopes());
    }

    [Fact]
    public void Load_UsesLeastPrivilegeDefaultsWhenNoAppSettingsExist()
    {
        var configuration = ExchangeConfigurationLoader.Load(_tempDirectory, _sharedConfigurationDirectory, _ => null);

        Assert.True(configuration.EnableGraphAfterExchangeConnect);
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
    }

    [Fact]
    public void Load_UsesPerMachineConfigurationOutsideInstallDirectory()
    {
        File.WriteAllText(
            Path.Combine(_sharedConfigurationDirectory, "appsettings.json"),
            """
            {
              "exchangeOnline": {
                "authenticationMode": "DeviceCode",
                "exchangeEnvironmentName": "O365Default",
                "graphScopes": [ "User.Read.All" ],
                "enableGraphAfterExchangeConnect": true
              }
            }
            """);

        var configuration = ExchangeConfigurationLoader.Load(_tempDirectory, _sharedConfigurationDirectory, _ => null);

        Assert.Equal(ExchangeAuthenticationMode.DeviceCode, configuration.AuthenticationMode);
        Assert.Equal(new[] { "User.Read.All" }, configuration.GraphScopes);
    }

    [Fact]
    public void Load_PerMachineConfigurationOverridesInstalledAppSettings()
    {
        File.WriteAllText(
            Path.Combine(_tempDirectory, "appsettings.json"),
            """
            {
              "exchangeOnline": {
                "authenticationMode": "Interactive",
                "exchangeEnvironmentName": "O365Default",
                "graphScopes": [ "Organization.Read.All" ],
                "enableGraphAfterExchangeConnect": true
              }
            }
            """);

        File.WriteAllText(
            Path.Combine(_sharedConfigurationDirectory, "appsettings.json"),
            """
            {
              "exchangeOnline": {
                "authenticationMode": "DeviceCode",
                "graphScopes": [ "Directory.Read.All", "User.Read.All" ]
              }
            }
            """);

        var configuration = ExchangeConfigurationLoader.Load(_tempDirectory, _sharedConfigurationDirectory, _ => null);

        Assert.Equal(ExchangeAuthenticationMode.DeviceCode, configuration.AuthenticationMode);
        Assert.Equal(new[] { "Directory.Read.All", "User.Read.All" }, configuration.GraphScopes);
    }

    [Fact]
    public void Load_UsesLegacyPerMachineConfigurationDirectory_WhenCurrentDirectoryIsEmpty()
    {
        var currentSharedDirectory = Path.Combine(_tempDirectory, "OnlyExo365");
        var legacySharedDirectory = Path.Combine(_tempDirectory, "ExchangeAdmin");
        Directory.CreateDirectory(currentSharedDirectory);
        Directory.CreateDirectory(legacySharedDirectory);

        File.WriteAllText(
            Path.Combine(legacySharedDirectory, "appsettings.json"),
            """
            {
              "exchangeOnline": {
                "authenticationMode": "DeviceCode",
                "exchangeEnvironmentName": "O365Default",
                "graphScopes": [ "User.Read.All" ],
                "enableGraphAfterExchangeConnect": true
              }
            }
            """);

        var configuration = ExchangeConfigurationLoader.Load(currentSharedDirectory, currentSharedDirectory, _ => null);

        Assert.Equal(ExchangeAuthenticationMode.DeviceCode, configuration.AuthenticationMode);
        Assert.Equal(new[] { "User.Read.All" }, configuration.GraphScopes);
    }

    [Fact]
    public void Load_ThrowsWhenPerMachineAppSettingsJsonIsMalformed()
    {
        File.WriteAllText(Path.Combine(_sharedConfigurationDirectory, "appsettings.json"), "{ invalid json");

        var exception = Assert.Throws<ExchangeConfigurationLoader.ExchangeConfigurationLoadException>(
            () => ExchangeConfigurationLoader.Load(_tempDirectory, _sharedConfigurationDirectory, _ => null));

        Assert.Contains("Invalid JSON", exception.Message);
        Assert.Contains("ProgramData", exception.FilePath);
    }

    [Fact]
    public void Load_ReportsEnvironmentVariableSourceForInvalidGraphTenantId()
    {
        var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase)
        {
            [ExchangeConfigurationEnvironmentVariables.GraphTenantId] = "tenant id with spaces"
        };

        var exception = Assert.Throws<ExchangeConfigurationLoader.ExchangeConfigurationLoadException>(
            () => ExchangeConfigurationLoader.Load(
                _tempDirectory,
                _sharedConfigurationDirectory,
                key => values.TryGetValue(key, out var value) ? value : null));

        Assert.Contains("GraphTenantId must be a tenant domain like 'contoso.onmicrosoft.com' or a tenant GUID.", exception.Message);
        Assert.Contains($"environment variable {ExchangeConfigurationEnvironmentVariables.GraphTenantId}", exception.Message);
    }

    [Fact]
    public void Load_ThrowsWhenEnvironmentAuthenticationModeIsInvalid()
    {
        var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase)
        {
            [ExchangeConfigurationEnvironmentVariables.AuthenticationMode] = "NotARealMode"
        };

        var exception = Assert.Throws<ExchangeConfigurationLoader.ExchangeConfigurationLoadException>(
            () => ExchangeConfigurationLoader.Load(
                _tempDirectory,
                _sharedConfigurationDirectory,
                key => values.TryGetValue(key, out var value) ? value : null));

        Assert.Contains("AuthenticationMode must be one of:", exception.Message);
        Assert.Contains($"environment variable {ExchangeConfigurationEnvironmentVariables.AuthenticationMode}", exception.Message);
    }

    [Fact]
    public void Load_ReportsFileSourceForInvalidExchangeOrganization()
    {
        File.WriteAllText(
            Path.Combine(_sharedConfigurationDirectory, "appsettings.json"),
            """
            {
              "exchangeOnline": {
                "exchangeOrganization": "tenant with spaces"
              }
            }
            """);

        var exception = Assert.Throws<ExchangeConfigurationLoader.ExchangeConfigurationLoadException>(
            () => ExchangeConfigurationLoader.Load(_tempDirectory, _sharedConfigurationDirectory, _ => null));

        Assert.Contains("ExchangeOrganization must be a tenant domain like 'contoso.onmicrosoft.com' or a tenant GUID.", exception.Message);
        Assert.Contains(Path.Combine(_sharedConfigurationDirectory, "appsettings.json"), exception.Message);
    }

    [Fact]
    public void Load_PartialEnvironmentOverridePreservesFileAuthenticationModeGraphScopesAndEnableGraphFlag()
    {
        File.WriteAllText(
            Path.Combine(_tempDirectory, "appsettings.json"),
            """
            {
              "exchangeOnline": {
                "authenticationMode": "AppCertificate",
                "applicationId": "11111111-1111-1111-1111-111111111111",
                "exchangeOrganization": "contoso.onmicrosoft.com",
                "certificateThumbprint": "thumbprint",
                "graphTenantId": "22222222-2222-2222-2222-222222222222",
                "graphScopes": [ "User.Read.All", "Organization.Read.All" ],
                "enableGraphAfterExchangeConnect": false
              }
            }
            """);

        var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase)
        {
            [ExchangeConfigurationEnvironmentVariables.ExchangeOrganization] = "fabrikam.onmicrosoft.com"
        };

        var configuration = ExchangeConfigurationLoader.Load(
            _tempDirectory,
            _sharedConfigurationDirectory,
            key => values.TryGetValue(key, out var value) ? value : null);

        Assert.Equal(ExchangeAuthenticationMode.AppCertificate, configuration.AuthenticationMode);
        Assert.Equal("fabrikam.onmicrosoft.com", configuration.ExchangeOrganization);
        Assert.Equal(new[] { "User.Read.All", "Organization.Read.All" }, configuration.GraphScopes);
        Assert.False(configuration.EnableGraphAfterExchangeConnect);
    }

    [Fact]
    public void Load_ReadsDefaultUsageLocationFromEnvironmentAndNormalizesIt()
    {
        var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase)
        {
            [ExchangeConfigurationEnvironmentVariables.DefaultUsageLocation] = " it "
        };

        var configuration = ExchangeConfigurationLoader.Load(
            _tempDirectory,
            _sharedConfigurationDirectory,
            key => values.TryGetValue(key, out var value) ? value : null);

        Assert.Equal("IT", configuration.DefaultUsageLocation);
    }

    [Fact]
    public void Load_ThrowsWhenEnvironmentDefaultUsageLocationIsInvalid()
    {
        var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase)
        {
            [ExchangeConfigurationEnvironmentVariables.DefaultUsageLocation] = "Italy"
        };

        var exception = Assert.Throws<ExchangeConfigurationLoader.ExchangeConfigurationLoadException>(
            () => ExchangeConfigurationLoader.Load(
                _tempDirectory,
                _sharedConfigurationDirectory,
                key => values.TryGetValue(key, out var value) ? value : null));

        Assert.Contains("DefaultUsageLocation must be a valid two-letter country/region code.", exception.Message);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDirectory))
            {
                Directory.Delete(_tempDirectory, recursive: true);
            }
        }
        catch
        {
        }
    }
}
