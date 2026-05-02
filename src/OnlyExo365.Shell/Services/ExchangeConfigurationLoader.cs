using System.IO;
using System.Text.Json;
using OnlyExo365.Contracts;

namespace OnlyExo365.Shell.Services;

public static class ExchangeConfigurationLoader
{
    private const string AppSettingsFileName = "appsettings.json";
    private const string VendorDirectoryName = "OnlyExo365";
    private const string ProductDirectoryName = "OnlyExo365";
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public static ExchangeOnlineConfiguration Load()
        => Load(
            AppDomain.CurrentDomain.BaseDirectory,
            GetPerMachineConfigurationDirectoryPath(),
            Environment.GetEnvironmentVariable);

    internal static ExchangeOnlineConfiguration Load(string baseDirectory, Func<string, string?> getEnvironmentVariable)
        => Load(baseDirectory, GetPerMachineConfigurationDirectoryPath(), getEnvironmentVariable);

    internal static ExchangeOnlineConfiguration Load(
        string baseDirectory,
        string sharedConfigurationDirectory,
        Func<string, string?> getEnvironmentVariable)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(baseDirectory);
        ArgumentException.ThrowIfNullOrWhiteSpace(sharedConfigurationDirectory);
        ArgumentNullException.ThrowIfNull(getEnvironmentVariable);

        var configuration = ExchangeOnlineConfiguration.CreateDefault();
        var valueSources = new ConfigurationValueSourceTracker();
        var installAppSettingsPath = Path.Combine(baseDirectory, AppSettingsFileName);
        var loadedConfigurationSource = ApplyFileConfiguration(configuration, installAppSettingsPath, valueSources);
        foreach (var sharedAppSettingsPath in EnumerateSharedAppSettingsPaths(sharedConfigurationDirectory))
        {
            var sharedConfigurationSource = ApplyFileConfiguration(configuration, sharedAppSettingsPath, valueSources);
            if (sharedConfigurationSource != null)
            {
                loadedConfigurationSource = sharedConfigurationSource;
            }
        }

        var environmentConfiguration = TryLoadEnvironmentOverrides(
            getEnvironmentVariable,
            valueSources,
            Path.Combine(sharedConfigurationDirectory, AppSettingsFileName));
        ApplyConfigurationOverlay(configuration, environmentConfiguration);
        configuration.GraphScopes = configuration.NormalizeGraphScopes().ToList();

        var validationErrors = configuration.Validate()
            .Select(error => AnnotateValidationError(error, valueSources))
            .ToList();
        if (validationErrors.Count > 0)
        {
            var message = loadedConfigurationSource != null
                ? $"Invalid Exchange configuration in '{loadedConfigurationSource}': {string.Join(" ", validationErrors)}"
                : $"Invalid Exchange configuration from runtime overrides: {string.Join(" ", validationErrors)}";

            throw new ExchangeConfigurationLoadException(
                message,
                loadedConfigurationSource ?? Path.Combine(sharedConfigurationDirectory, AppSettingsFileName));
        }

        return configuration;
    }

    internal static string GetPerMachineConfigurationDirectoryPath()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            VendorDirectoryName,
            ProductDirectoryName);
    }

    private static string? ApplyFileConfiguration(
        ExchangeOnlineConfiguration configuration,
        string filePath,
        ConfigurationValueSourceTracker valueSources)
    {
        var fileConfiguration = TryLoadFromAppSettings(filePath);
        if (fileConfiguration == null)
        {
            return null;
        }

        valueSources.TrackFileConfiguration(fileConfiguration.Configuration, filePath);
        ApplyConfigurationOverlay(configuration, fileConfiguration);
        return filePath;
    }

    private static ConfigurationOverlay? TryLoadFromAppSettings(string filePath)
    {
        if (!File.Exists(filePath))
        {
            return null;
        }

        try
        {
            var json = File.ReadAllText(filePath);
            using var document = JsonDocument.Parse(json);
            if (!TryGetPropertyIgnoreCase(document.RootElement, "exchangeOnline", out var exchangeOnlineElement) ||
                exchangeOnlineElement.ValueKind == JsonValueKind.Null ||
                exchangeOnlineElement.ValueKind != JsonValueKind.Object)
            {
                return null;
            }

            var configuration = JsonSerializer.Deserialize<ExchangeOnlineConfiguration>(
                exchangeOnlineElement.GetRawText(),
                JsonOptions);
            if (configuration == null)
            {
                return null;
            }

            return ConfigurationOverlay.FromJsonObject(configuration, exchangeOnlineElement);
        }
        catch (JsonException ex)
        {
            throw new ExchangeConfigurationLoadException(
                $"Invalid JSON in '{filePath}' at line {ex.LineNumber}, byte position {ex.BytePositionInLine}.",
                filePath,
                ex);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw new ExchangeConfigurationLoadException(
                $"Unable to read Exchange configuration file '{filePath}': {ex.Message}",
                filePath,
                ex);
        }
        catch
        {
            throw;
        }
    }

    private static ConfigurationOverlay? TryLoadEnvironmentOverrides(
        Func<string, string?> getEnvironmentVariable,
        ConfigurationValueSourceTracker valueSources,
        string configurationErrorPath)
    {
        var hasAnyOverride = false;
        var configuration = new ExchangeOnlineConfiguration();
        var overlay = new ConfigurationOverlay(configuration);

        var environmentName = getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.ExchangeEnvironmentName);
        if (!string.IsNullOrWhiteSpace(environmentName))
        {
            configuration.ExchangeEnvironmentName = environmentName.Trim();
            overlay.HasExchangeEnvironmentName = true;
            hasAnyOverride = true;
        }

        var authMode = getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.AuthenticationMode);
        if (!string.IsNullOrWhiteSpace(authMode))
        {
            try
            {
                configuration.AuthenticationMode = ExchangeOnlineConfiguration.ParseAuthenticationModeOrThrow(authMode);
                overlay.HasAuthenticationMode = true;
                hasAnyOverride = true;
            }
            catch (InvalidOperationException ex)
            {
                throw new ExchangeConfigurationLoadException(
                    $"{ex.Message} Source: environment variable {ExchangeConfigurationEnvironmentVariables.AuthenticationMode}.",
                    configurationErrorPath,
                    ex);
            }
        }

        hasAnyOverride |= TrySetOptional(
            getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.ExchangeOrganization),
            value =>
            {
                configuration.ExchangeOrganization = value;
                overlay.HasExchangeOrganization = true;
                valueSources.ExchangeOrganizationSource = $"environment variable {ExchangeConfigurationEnvironmentVariables.ExchangeOrganization}";
            });
        hasAnyOverride |= TrySetOptional(getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.DelegatedOrganization), value =>
        {
            configuration.DelegatedOrganization = value;
            overlay.HasDelegatedOrganization = true;
        });
        hasAnyOverride |= TrySetOptional(getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.UserPrincipalNameHint), value =>
        {
            configuration.UserPrincipalNameHint = value;
            overlay.HasUserPrincipalNameHint = true;
        });
        hasAnyOverride |= TrySetOptional(getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.ApplicationId), value =>
        {
            configuration.ApplicationId = value;
            overlay.HasApplicationId = true;
        });
        hasAnyOverride |= TrySetOptional(getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.CertificateThumbprint), value =>
        {
            configuration.CertificateThumbprint = value;
            overlay.HasCertificateThumbprint = true;
        });
        hasAnyOverride |= TrySetOptional(getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.CertificateSubjectName), value =>
        {
            configuration.CertificateSubjectName = value;
            overlay.HasCertificateSubjectName = true;
        });
        hasAnyOverride |= TrySetOptional(getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.ManagedIdentityAccountId), value =>
        {
            configuration.ManagedIdentityAccountId = value;
            overlay.HasManagedIdentityAccountId = true;
        });
        hasAnyOverride |= TrySetOptional(
            getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.GraphTenantId),
            value =>
            {
                configuration.GraphTenantId = value;
                overlay.HasGraphTenantId = true;
                valueSources.GraphTenantIdSource = $"environment variable {ExchangeConfigurationEnvironmentVariables.GraphTenantId}";
            });

        var graphScopes = getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.GraphScopes);
        if (!string.IsNullOrWhiteSpace(graphScopes))
        {
            configuration.GraphScopes = graphScopes
                .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .ToList();
            overlay.HasGraphScopes = true;
            hasAnyOverride = true;
        }

        var graphLicenseWriteScopes = getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.GraphLicenseWriteScopes);
        if (!string.IsNullOrWhiteSpace(graphLicenseWriteScopes))
        {
            configuration.GraphLicenseWriteScopes = graphLicenseWriteScopes
                .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .ToList();
            overlay.HasGraphLicenseWriteScopes = true;
            hasAnyOverride = true;
        }

        var defaultUsageLocation = getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.DefaultUsageLocation);
        if (!string.IsNullOrWhiteSpace(defaultUsageLocation))
        {
            configuration.DefaultUsageLocation = ExchangeOnlineConfiguration.NormalizeUsageLocation(defaultUsageLocation);
            overlay.HasDefaultUsageLocation = true;
            hasAnyOverride = true;
        }

        var enableGraph = getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.EnableGraphAfterExchangeConnect);
        if (!string.IsNullOrWhiteSpace(enableGraph))
        {
            try
            {
                configuration.EnableGraphAfterExchangeConnect = ExchangeOnlineConfiguration.ParseBooleanOrThrow(
                    enableGraph,
                    ExchangeConfigurationEnvironmentVariables.EnableGraphAfterExchangeConnect);
                overlay.HasEnableGraphAfterExchangeConnect = true;
                hasAnyOverride = true;
            }
            catch (InvalidOperationException ex)
            {
                throw new ExchangeConfigurationLoadException(
                    $"{ex.Message} Source: environment variable {ExchangeConfigurationEnvironmentVariables.EnableGraphAfterExchangeConnect}.",
                    configurationErrorPath,
                    ex);
            }
        }

        return hasAnyOverride ? overlay : null;
    }

    private static bool TrySetOptional(string? value, Action<string> assign)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        assign(value.Trim());
        return true;
    }

    private static bool TryGetPropertyIgnoreCase(JsonElement element, string propertyName, out JsonElement propertyValue)
    {
        foreach (var property in element.EnumerateObject())
        {
            if (property.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase))
            {
                propertyValue = property.Value;
                return true;
            }
        }

        propertyValue = default;
        return false;
    }

    private static IEnumerable<string> EnumerateSharedAppSettingsPaths(string sharedConfigurationDirectory)
    {
        foreach (var directory in EnumerateSharedConfigurationDirectories(sharedConfigurationDirectory))
        {
            yield return Path.Combine(directory, AppSettingsFileName);
        }
    }

    private static IEnumerable<string> EnumerateSharedConfigurationDirectories(string sharedConfigurationDirectory)
    {
        if (string.IsNullOrWhiteSpace(sharedConfigurationDirectory))
        {
            yield break;
        }

        yield return sharedConfigurationDirectory;
    }

    private static string? NormalizeOptional(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    private static void ApplyConfigurationOverlay(ExchangeOnlineConfiguration configuration, ConfigurationOverlay? overlay)
    {
        if (overlay == null)
        {
            return;
        }

        var overrides = overlay.Configuration;
        if (overlay.HasAuthenticationMode)
        {
            configuration.AuthenticationMode = overrides.AuthenticationMode;
        }

        if (overlay.HasExchangeEnvironmentName && !string.IsNullOrWhiteSpace(overrides.ExchangeEnvironmentName))
        {
            configuration.ExchangeEnvironmentName = overrides.ExchangeEnvironmentName.Trim();
        }

        if (overlay.HasExchangeOrganization)
        {
            configuration.ExchangeOrganization = NormalizeOptional(overrides.ExchangeOrganization) ?? configuration.ExchangeOrganization;
        }

        if (overlay.HasDelegatedOrganization)
        {
            configuration.DelegatedOrganization = NormalizeOptional(overrides.DelegatedOrganization) ?? configuration.DelegatedOrganization;
        }

        if (overlay.HasUserPrincipalNameHint)
        {
            configuration.UserPrincipalNameHint = NormalizeOptional(overrides.UserPrincipalNameHint) ?? configuration.UserPrincipalNameHint;
        }

        if (overlay.HasApplicationId)
        {
            configuration.ApplicationId = NormalizeOptional(overrides.ApplicationId) ?? configuration.ApplicationId;
        }

        if (overlay.HasCertificateThumbprint)
        {
            configuration.CertificateThumbprint = NormalizeOptional(overrides.CertificateThumbprint) ?? configuration.CertificateThumbprint;
        }

        if (overlay.HasCertificateSubjectName)
        {
            configuration.CertificateSubjectName = NormalizeOptional(overrides.CertificateSubjectName) ?? configuration.CertificateSubjectName;
        }

        if (overlay.HasManagedIdentityAccountId)
        {
            configuration.ManagedIdentityAccountId = NormalizeOptional(overrides.ManagedIdentityAccountId) ?? configuration.ManagedIdentityAccountId;
        }

        if (overlay.HasGraphTenantId)
        {
            configuration.GraphTenantId = NormalizeOptional(overrides.GraphTenantId) ?? configuration.GraphTenantId;
        }

        if (overlay.HasGraphScopes)
        {
            configuration.GraphScopes = configuration.NormalizeGraphScopes(overrides.GraphScopes).ToList();
        }

        if (overlay.HasGraphLicenseWriteScopes)
        {
            configuration.GraphLicenseWriteScopes = configuration.NormalizeGraphLicenseWriteScopes(overrides.GraphLicenseWriteScopes).ToList();
        }

        if (overlay.HasDefaultUsageLocation)
        {
            configuration.DefaultUsageLocation = ExchangeOnlineConfiguration.NormalizeUsageLocation(overrides.DefaultUsageLocation) ?? configuration.DefaultUsageLocation;
        }

        if (overlay.HasEnableGraphAfterExchangeConnect)
        {
            configuration.EnableGraphAfterExchangeConnect = overrides.EnableGraphAfterExchangeConnect;
        }
    }

    private static string AnnotateValidationError(string error, ConfigurationValueSourceTracker valueSources)
    {
        if (error.StartsWith("ExchangeOrganization ", StringComparison.Ordinal))
        {
            return AppendSource(error, valueSources.ExchangeOrganizationSource);
        }

        if (error.StartsWith("GraphTenantId ", StringComparison.Ordinal))
        {
            return AppendSource(error, valueSources.GraphTenantIdSource);
        }

        return error;
    }

    private static string AppendSource(string error, string? source)
        => string.IsNullOrWhiteSpace(source)
            ? error
            : $"{error} Source: {source}.";

    private sealed class ConfigurationValueSourceTracker
    {
        public string? ExchangeOrganizationSource { get; set; }

        public string? GraphTenantIdSource { get; set; }

        public void TrackFileConfiguration(ExchangeOnlineConfiguration configuration, string filePath)
        {
            if (!string.IsNullOrWhiteSpace(configuration.ExchangeOrganization))
            {
                ExchangeOrganizationSource = filePath;
            }

            if (!string.IsNullOrWhiteSpace(configuration.GraphTenantId))
            {
                GraphTenantIdSource = filePath;
            }
        }
    }

    private sealed class ConfigurationOverlay
    {
        public ConfigurationOverlay(ExchangeOnlineConfiguration configuration)
        {
            Configuration = configuration;
        }

        public ExchangeOnlineConfiguration Configuration { get; }

        public bool HasAuthenticationMode { get; set; }
        public bool HasExchangeEnvironmentName { get; set; }
        public bool HasExchangeOrganization { get; set; }
        public bool HasDelegatedOrganization { get; set; }
        public bool HasUserPrincipalNameHint { get; set; }
        public bool HasApplicationId { get; set; }
        public bool HasCertificateThumbprint { get; set; }
        public bool HasCertificateSubjectName { get; set; }
        public bool HasManagedIdentityAccountId { get; set; }
        public bool HasGraphTenantId { get; set; }
        public bool HasGraphScopes { get; set; }
        public bool HasGraphLicenseWriteScopes { get; set; }
        public bool HasDefaultUsageLocation { get; set; }
        public bool HasEnableGraphAfterExchangeConnect { get; set; }

        public static ConfigurationOverlay FromJsonObject(ExchangeOnlineConfiguration configuration, JsonElement exchangeOnlineElement)
        {
            var overlay = new ConfigurationOverlay(configuration);
            foreach (var property in exchangeOnlineElement.EnumerateObject())
            {
                switch (property.Name)
                {
                    case var name when name.Equals("authenticationMode", StringComparison.OrdinalIgnoreCase):
                        overlay.HasAuthenticationMode = true;
                        break;
                    case var name when name.Equals("exchangeEnvironmentName", StringComparison.OrdinalIgnoreCase):
                        overlay.HasExchangeEnvironmentName = true;
                        break;
                    case var name when name.Equals("exchangeOrganization", StringComparison.OrdinalIgnoreCase):
                        overlay.HasExchangeOrganization = true;
                        break;
                    case var name when name.Equals("delegatedOrganization", StringComparison.OrdinalIgnoreCase):
                        overlay.HasDelegatedOrganization = true;
                        break;
                    case var name when name.Equals("userPrincipalNameHint", StringComparison.OrdinalIgnoreCase):
                        overlay.HasUserPrincipalNameHint = true;
                        break;
                    case var name when name.Equals("applicationId", StringComparison.OrdinalIgnoreCase):
                        overlay.HasApplicationId = true;
                        break;
                    case var name when name.Equals("certificateThumbprint", StringComparison.OrdinalIgnoreCase):
                        overlay.HasCertificateThumbprint = true;
                        break;
                    case var name when name.Equals("certificateSubjectName", StringComparison.OrdinalIgnoreCase):
                        overlay.HasCertificateSubjectName = true;
                        break;
                    case var name when name.Equals("managedIdentityAccountId", StringComparison.OrdinalIgnoreCase):
                        overlay.HasManagedIdentityAccountId = true;
                        break;
                    case var name when name.Equals("graphTenantId", StringComparison.OrdinalIgnoreCase):
                        overlay.HasGraphTenantId = true;
                        break;
                    case var name when name.Equals("graphScopes", StringComparison.OrdinalIgnoreCase):
                        overlay.HasGraphScopes = true;
                        break;
                    case var name when name.Equals("graphLicenseWriteScopes", StringComparison.OrdinalIgnoreCase):
                        overlay.HasGraphLicenseWriteScopes = true;
                        break;
                    case var name when name.Equals("defaultUsageLocation", StringComparison.OrdinalIgnoreCase):
                        overlay.HasDefaultUsageLocation = true;
                        break;
                    case var name when name.Equals("enableGraphAfterExchangeConnect", StringComparison.OrdinalIgnoreCase):
                        overlay.HasEnableGraphAfterExchangeConnect = true;
                        break;
                }
            }

            return overlay;
        }
    }

    // -------------------------------------------------------------------------
    // License Catalog configuration
    // -------------------------------------------------------------------------

    /// <summary>
    /// Loads the <c>licensingCatalog</c> section from the standard
    /// appsettings resolution chain (install dir, then ProgramData shared
    /// config).  Never throws — configuration errors produce safe defaults
    /// so a missing or malformed section never blocks startup.
    /// </summary>
    public static LicenseCatalogConfiguration LoadLicenseCatalogConfiguration()
        => LoadLicenseCatalogConfiguration(
            AppDomain.CurrentDomain.BaseDirectory,
            GetPerMachineConfigurationDirectoryPath());

    private static readonly JsonSerializerOptions CatalogJsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        Converters = { new System.Text.Json.Serialization.JsonStringEnumConverter() }
    };

    internal static LicenseCatalogConfiguration LoadLicenseCatalogConfiguration(
        string baseDirectory,
        string sharedConfigurationDirectory)
    {
        var config = LicenseCatalogConfiguration.CreateDefault();

        // Apply install-directory appsettings first, then ProgramData override.
        foreach (var dir in new[] { baseDirectory }.Concat(EnumerateSharedConfigurationDirectories(sharedConfigurationDirectory)))
        {
            var path = Path.Combine(dir, AppSettingsFileName);
            if (!File.Exists(path))
            {
                continue;
            }

            try
            {
                var json = File.ReadAllText(path);
                using var document = JsonDocument.Parse(json);
                if (!TryGetPropertyIgnoreCase(document.RootElement,
                        LicenseCatalogConfiguration.SectionName, out var section) ||
                    section.ValueKind != JsonValueKind.Object)
                {
                    continue;
                }

                var overlay = JsonSerializer.Deserialize<LicenseCatalogConfiguration>(
                    section.GetRawText(), CatalogJsonOptions);

                if (overlay == null)
                {
                    continue;
                }

                // Merge non-default values from the overlay.
                config.AutoUpdateMode = overlay.AutoUpdateMode;
                config.CheckOnStartup = overlay.CheckOnStartup;

                if (!string.IsNullOrWhiteSpace(overlay.RemoteSource))
                {
                    config.RemoteSource = overlay.RemoteSource;
                }

                if (overlay.DownloadTimeoutSeconds > 0)
                {
                    config.DownloadTimeoutSeconds = overlay.DownloadTimeoutSeconds;
                }

                if (!string.IsNullOrWhiteSpace(overlay.LocalCachePath))
                {
                    config.LocalCachePath = overlay.LocalCachePath;
                }
            }
            catch
            {
                // Catalog config errors are non-fatal; continue with defaults.
            }
        }

        return config;
    }

    internal sealed class ExchangeConfigurationLoadException : InvalidOperationException
    {
        public ExchangeConfigurationLoadException(string message, string filePath, Exception? innerException = null)
            : base(message, innerException)
        {
            FilePath = filePath;
        }

        public string FilePath { get; }
    }
}

