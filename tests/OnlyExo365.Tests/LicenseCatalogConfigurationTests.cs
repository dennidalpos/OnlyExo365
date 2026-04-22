using System.IO;
using OnlyExo365.Shell.Services;

namespace OnlyExo365.Tests;

public sealed class LicenseCatalogConfigurationTests : IDisposable
{
    private readonly string _tempDir = Path.Combine(
        Path.GetTempPath(), "OnlyExo365.Tests", Guid.NewGuid().ToString("N"));

    private readonly string _sharedDir;

    public LicenseCatalogConfigurationTests()
    {
        Directory.CreateDirectory(_tempDir);
        _sharedDir = Path.Combine(_tempDir, "ProgramData");
        Directory.CreateDirectory(_sharedDir);
    }

    [Fact]
    public void LoadLicenseCatalogConfiguration_ReturnsDefaults_WhenSectionAbsent()
    {
        // appsettings.json with no licensingCatalog section.
        File.WriteAllText(
            Path.Combine(_tempDir, "appsettings.json"),
            """{ "ExchangeOnline": {} }""");

        var config = ExchangeConfigurationLoader.LoadLicenseCatalogConfiguration(
            _tempDir, _sharedDir);

        Assert.Equal(CatalogAutoUpdateMode.Daily, config.AutoUpdateMode);
        Assert.True(config.CheckOnStartup);
        Assert.Equal(LicenseCatalogConfiguration.DefaultRemoteSource, config.RemoteSource);
        Assert.Equal(30, config.DownloadTimeoutSeconds);
        Assert.Null(config.LocalCachePath);
    }

    [Fact]
    public void LoadLicenseCatalogConfiguration_ParsesAllFields_WhenSectionPresent()
    {
        File.WriteAllText(
            Path.Combine(_tempDir, "appsettings.json"),
            """
            {
              "licensingCatalog": {
                "autoUpdateMode": "Monthly",
                "checkOnStartup": false,
                "remoteSource": "https://example.com/sku",
                "downloadTimeoutSeconds": 60,
                "localCachePath": "C:\\Temp\\catalog"
              }
            }
            """);

        var config = ExchangeConfigurationLoader.LoadLicenseCatalogConfiguration(
            _tempDir, _sharedDir);

        Assert.Equal(CatalogAutoUpdateMode.Monthly, config.AutoUpdateMode);
        Assert.False(config.CheckOnStartup);
        Assert.Equal("https://example.com/sku", config.RemoteSource);
        Assert.Equal(60, config.DownloadTimeoutSeconds);
        Assert.Equal("C:\\Temp\\catalog", config.LocalCachePath);
    }

    [Fact]
    public void LoadLicenseCatalogConfiguration_ReturnsDefaults_OnMalformedJson()
    {
        File.WriteAllText(Path.Combine(_tempDir, "appsettings.json"), "{ invalid json");

        // Must not throw even though the JSON is malformed.
        var config = ExchangeConfigurationLoader.LoadLicenseCatalogConfiguration(
            _tempDir, _sharedDir);

        Assert.Equal(CatalogAutoUpdateMode.Daily, config.AutoUpdateMode);
        Assert.True(config.CheckOnStartup);
    }

    [Fact]
    public void LoadLicenseCatalogConfiguration_SharedDirOverridesInstallDir()
    {
        File.WriteAllText(
            Path.Combine(_tempDir, "appsettings.json"),
            """
            {
              "licensingCatalog": {
                "autoUpdateMode": "Daily"
              }
            }
            """);

        File.WriteAllText(
            Path.Combine(_sharedDir, "appsettings.json"),
            """
            {
              "licensingCatalog": {
                "autoUpdateMode": "Disabled"
              }
            }
            """);

        var config = ExchangeConfigurationLoader.LoadLicenseCatalogConfiguration(
            _tempDir, _sharedDir);

        // Shared dir wins.
        Assert.Equal(CatalogAutoUpdateMode.Disabled, config.AutoUpdateMode);
    }

    [Fact]
    public void ResolveLocalCachePath_UsesLocalAppData_WhenNullConfigured()
    {
        var config = new LicenseCatalogConfiguration { LocalCachePath = null };
        var path = config.ResolveLocalCachePath();

        Assert.Contains("OnlyExo365", path);
        Assert.Contains("LicenseCatalog", path);
    }

    [Fact]
    public void ResolveLocalCachePath_UsesProvidedPath_WhenConfigured()
    {
        const string custom = "D:\\Custom\\Path";
        var config = new LicenseCatalogConfiguration { LocalCachePath = custom };

        Assert.Equal(custom, config.ResolveLocalCachePath());
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_tempDir, recursive: true);
        }
        catch
        {
            // Best-effort cleanup.
        }
    }
}

