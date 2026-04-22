using System.Text.Json.Serialization;

namespace OnlyExo365.Contracts;

public enum ExchangeAuthenticationMode
{
    Interactive,
    DeviceCode,
    AppCertificate,
    ManagedIdentity
}

public static class ExchangeConfigurationEnvironmentVariables
{
    public const string ExchangeEnvironmentName = "ONLYEXO365_EXO_ENV";
    public const string AuthenticationMode = "ONLYEXO365_AUTH_MODE";
    public const string ExchangeOrganization = "ONLYEXO365_EXO_ORGANIZATION";
    public const string DelegatedOrganization = "ONLYEXO365_EXO_DELEGATED_ORGANIZATION";
    public const string UserPrincipalNameHint = "ONLYEXO365_EXO_UPN_HINT";
    public const string ApplicationId = "ONLYEXO365_APP_ID";
    public const string CertificateThumbprint = "ONLYEXO365_CERT_THUMBPRINT";
    public const string CertificateSubjectName = "ONLYEXO365_CERT_SUBJECT";
    public const string ManagedIdentityAccountId = "ONLYEXO365_MANAGED_IDENTITY_ACCOUNT_ID";
    public const string GraphTenantId = "ONLYEXO365_GRAPH_TENANT_ID";
    public const string GraphScopes = "ONLYEXO365_GRAPH_SCOPES";
    public const string GraphLicenseWriteScopes = "ONLYEXO365_GRAPH_LICENSE_WRITE_SCOPES";
    public const string DefaultUsageLocation = "ONLYEXO365_DEFAULT_USAGE_LOCATION";
    public const string EnableGraphAfterExchangeConnect = "ONLYEXO365_ENABLE_GRAPH";
}

public sealed class ExchangeOnlineConfiguration
{
    private static readonly string[] DefaultGraphScopes =
    {
        "Organization.Read.All",
        "Directory.Read.All",
        "RoleManagement.Read.Directory",
        "User.Read.All"
    };

    [JsonPropertyName("authenticationMode")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public ExchangeAuthenticationMode AuthenticationMode { get; set; } = ExchangeAuthenticationMode.Interactive;

    [JsonPropertyName("exchangeEnvironmentName")]
    public string ExchangeEnvironmentName { get; set; } = "O365Default";

    [JsonPropertyName("exchangeOrganization")]
    public string? ExchangeOrganization { get; set; }

    [JsonPropertyName("delegatedOrganization")]
    public string? DelegatedOrganization { get; set; }

    [JsonPropertyName("userPrincipalNameHint")]
    public string? UserPrincipalNameHint { get; set; }

    [JsonPropertyName("applicationId")]
    public string? ApplicationId { get; set; }

    [JsonPropertyName("certificateThumbprint")]
    public string? CertificateThumbprint { get; set; }

    [JsonPropertyName("certificateSubjectName")]
    public string? CertificateSubjectName { get; set; }

    [JsonPropertyName("managedIdentityAccountId")]
    public string? ManagedIdentityAccountId { get; set; }

    [JsonPropertyName("graphTenantId")]
    public string? GraphTenantId { get; set; }

    [JsonPropertyName("graphScopes")]
    public List<string> GraphScopes { get; set; } = DefaultGraphScopes.ToList();

    [JsonPropertyName("graphLicenseWriteScopes")]
    public List<string>? GraphLicenseWriteScopes { get; set; }

    [JsonPropertyName("defaultUsageLocation")]
    public string? DefaultUsageLocation { get; set; }

    [JsonPropertyName("enableGraphAfterExchangeConnect")]
    public bool EnableGraphAfterExchangeConnect { get; set; } = true;

    public static ExchangeOnlineConfiguration CreateDefault() => new();

    public ExchangeOnlineConfiguration Clone()
    {
        return new ExchangeOnlineConfiguration
        {
            AuthenticationMode = AuthenticationMode,
            ExchangeEnvironmentName = ExchangeEnvironmentName,
            ExchangeOrganization = ExchangeOrganization,
            DelegatedOrganization = DelegatedOrganization,
            UserPrincipalNameHint = UserPrincipalNameHint,
            ApplicationId = ApplicationId,
            CertificateThumbprint = CertificateThumbprint,
            CertificateSubjectName = CertificateSubjectName,
            ManagedIdentityAccountId = ManagedIdentityAccountId,
            GraphTenantId = GraphTenantId,
            GraphScopes = NormalizeGraphScopes().ToList(),
            GraphLicenseWriteScopes = GraphLicenseWriteScopes is null
                ? null
                : NormalizeGraphLicenseWriteScopes().ToList(),
            DefaultUsageLocation = NormalizeUsageLocation(DefaultUsageLocation),
            EnableGraphAfterExchangeConnect = EnableGraphAfterExchangeConnect
        };
    }

    public void ApplyOverrides(ExchangeOnlineConfiguration? overrides)
    {
        if (overrides == null)
        {
            return;
        }

        AuthenticationMode = overrides.AuthenticationMode;

        if (!string.IsNullOrWhiteSpace(overrides.ExchangeEnvironmentName))
        {
            ExchangeEnvironmentName = overrides.ExchangeEnvironmentName.Trim();
        }

        ExchangeOrganization = NormalizeOptional(overrides.ExchangeOrganization) ?? ExchangeOrganization;
        DelegatedOrganization = NormalizeOptional(overrides.DelegatedOrganization) ?? DelegatedOrganization;
        UserPrincipalNameHint = NormalizeOptional(overrides.UserPrincipalNameHint) ?? UserPrincipalNameHint;
        ApplicationId = NormalizeOptional(overrides.ApplicationId) ?? ApplicationId;
        CertificateThumbprint = NormalizeOptional(overrides.CertificateThumbprint) ?? CertificateThumbprint;
        CertificateSubjectName = NormalizeOptional(overrides.CertificateSubjectName) ?? CertificateSubjectName;
        ManagedIdentityAccountId = NormalizeOptional(overrides.ManagedIdentityAccountId) ?? ManagedIdentityAccountId;
        GraphTenantId = NormalizeOptional(overrides.GraphTenantId) ?? GraphTenantId;
        DefaultUsageLocation = NormalizeUsageLocation(overrides.DefaultUsageLocation) ?? DefaultUsageLocation;

        var normalizedScopes = overrides.NormalizeGraphScopes();
        if (normalizedScopes.Count > 0)
        {
            GraphScopes = normalizedScopes.ToList();
        }

        if (overrides.GraphLicenseWriteScopes is not null)
        {
            GraphLicenseWriteScopes = overrides.NormalizeGraphLicenseWriteScopes().ToList();
        }

        EnableGraphAfterExchangeConnect = overrides.EnableGraphAfterExchangeConnect;
    }

    public IReadOnlyList<string> NormalizeGraphScopes(IEnumerable<string>? scopes = null)
    {
        var normalizedScopes = (scopes ?? GraphScopes ?? new List<string>())
            .Select(scope => scope?.Trim())
            .Where(scope => !string.IsNullOrWhiteSpace(scope))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Cast<string>()
            .ToList();

        return normalizedScopes.Count == 0 ? DefaultGraphScopes.ToList() : normalizedScopes;
    }

    public IReadOnlyList<string> NormalizeGraphLicenseWriteScopes(IEnumerable<string>? scopes = null)
    {
        return (scopes ?? GraphLicenseWriteScopes ?? new List<string>())
            .Select(scope => scope?.Trim())
            .Where(scope => !string.IsNullOrWhiteSpace(scope))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Cast<string>()
            .ToList();
    }

    public IReadOnlyList<string> GetGraphScopesForLicenseWrite()
    {
        return NormalizeGraphScopes()
            .Concat(NormalizeGraphLicenseWriteScopes())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public IReadOnlyList<string> Validate()
    {
        var errors = new List<string>();

        if (string.IsNullOrWhiteSpace(ExchangeEnvironmentName))
        {
            errors.Add("ExchangeEnvironmentName is required.");
        }

        if (!Enum.IsDefined(AuthenticationMode))
        {
            errors.Add("AuthenticationMode is invalid.");
        }

        var exchangeOrganizationError = ValidateExchangeOrganizationValue(ExchangeOrganization);
        if (exchangeOrganizationError is not null)
        {
            errors.Add(exchangeOrganizationError);
        }

        var graphTenantIdError = ValidateGraphTenantIdValue(GraphTenantId);
        if (graphTenantIdError is not null)
        {
            errors.Add(graphTenantIdError);
        }

        var defaultUsageLocationError = ValidateUsageLocationValue(DefaultUsageLocation, "DefaultUsageLocation");
        if (defaultUsageLocationError is not null)
        {
            errors.Add(defaultUsageLocationError);
        }

        switch (AuthenticationMode)
        {
            case ExchangeAuthenticationMode.DeviceCode:
                if (!string.IsNullOrWhiteSpace(CertificateThumbprint) || !string.IsNullOrWhiteSpace(CertificateSubjectName))
                {
                    errors.Add("Certificate settings are not applicable to DeviceCode authentication.");
                }
                break;

            case ExchangeAuthenticationMode.AppCertificate:
                if (string.IsNullOrWhiteSpace(ApplicationId))
                {
                    errors.Add("ApplicationId is required for AppCertificate authentication.");
                }

                if (string.IsNullOrWhiteSpace(ExchangeOrganization))
                {
                    errors.Add("ExchangeOrganization is required for AppCertificate authentication.");
                }

                if (string.IsNullOrWhiteSpace(CertificateThumbprint) && string.IsNullOrWhiteSpace(CertificateSubjectName))
                {
                    errors.Add("CertificateThumbprint or CertificateSubjectName is required for AppCertificate authentication.");
                }

                if (EnableGraphAfterExchangeConnect && string.IsNullOrWhiteSpace(GraphTenantId))
                {
                    errors.Add("GraphTenantId is required when Graph is enabled with AppCertificate authentication.");
                }
                break;

            case ExchangeAuthenticationMode.ManagedIdentity:
                if (string.IsNullOrWhiteSpace(ExchangeOrganization))
                {
                    errors.Add("ExchangeOrganization is required for ManagedIdentity authentication.");
                }

                if (!string.IsNullOrWhiteSpace(CertificateThumbprint) || !string.IsNullOrWhiteSpace(CertificateSubjectName))
                {
                    errors.Add("Certificate settings are not applicable to ManagedIdentity authentication.");
                }
                break;
        }

        return errors;
    }

    public static void ThrowIfInvalidExchangeOrganization(string? value)
    {
        var error = ValidateExchangeOrganizationValue(value);
        if (error is not null)
        {
            throw new InvalidOperationException(error);
        }
    }

    public static void ThrowIfInvalidGraphTenantId(string? value)
    {
        var error = ValidateGraphTenantIdValue(value);
        if (error is not null)
        {
            throw new InvalidOperationException(error);
        }
    }

    public void ApplyEnvironmentVariables(System.Collections.Specialized.StringDictionary environmentVariables)
    {
        ArgumentNullException.ThrowIfNull(environmentVariables);

        environmentVariables[ExchangeConfigurationEnvironmentVariables.ExchangeEnvironmentName] = ExchangeEnvironmentName;
        environmentVariables[ExchangeConfigurationEnvironmentVariables.AuthenticationMode] = AuthenticationMode.ToString();
        environmentVariables[ExchangeConfigurationEnvironmentVariables.ExchangeOrganization] = ExchangeOrganization;
        environmentVariables[ExchangeConfigurationEnvironmentVariables.DelegatedOrganization] = DelegatedOrganization;
        environmentVariables[ExchangeConfigurationEnvironmentVariables.UserPrincipalNameHint] = UserPrincipalNameHint;
        environmentVariables[ExchangeConfigurationEnvironmentVariables.ApplicationId] = ApplicationId;
        environmentVariables[ExchangeConfigurationEnvironmentVariables.CertificateThumbprint] = CertificateThumbprint;
        environmentVariables[ExchangeConfigurationEnvironmentVariables.CertificateSubjectName] = CertificateSubjectName;
        environmentVariables[ExchangeConfigurationEnvironmentVariables.ManagedIdentityAccountId] = ManagedIdentityAccountId;
        environmentVariables[ExchangeConfigurationEnvironmentVariables.GraphTenantId] = GraphTenantId;
        environmentVariables[ExchangeConfigurationEnvironmentVariables.GraphScopes] = string.Join(";", NormalizeGraphScopes());
        environmentVariables[ExchangeConfigurationEnvironmentVariables.GraphLicenseWriteScopes] = string.Join(";", NormalizeGraphLicenseWriteScopes());
        environmentVariables[ExchangeConfigurationEnvironmentVariables.DefaultUsageLocation] = NormalizeUsageLocation(DefaultUsageLocation);
        environmentVariables[ExchangeConfigurationEnvironmentVariables.EnableGraphAfterExchangeConnect] = EnableGraphAfterExchangeConnect ? "1" : "0";
    }

    public static ExchangeOnlineConfiguration FromEnvironmentVariables(Func<string, string?> getEnvironmentVariable)
    {
        ArgumentNullException.ThrowIfNull(getEnvironmentVariable);

        var configuration = CreateDefault();

        configuration.ExchangeEnvironmentName = ReadOrDefault(
            getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.ExchangeEnvironmentName),
            configuration.ExchangeEnvironmentName);

        configuration.ExchangeOrganization = ReadOptional(getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.ExchangeOrganization));
        configuration.DelegatedOrganization = ReadOptional(getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.DelegatedOrganization));
        configuration.UserPrincipalNameHint = ReadOptional(getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.UserPrincipalNameHint));
        configuration.ApplicationId = ReadOptional(getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.ApplicationId));
        configuration.CertificateThumbprint = ReadOptional(getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.CertificateThumbprint));
        configuration.CertificateSubjectName = ReadOptional(getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.CertificateSubjectName));
        configuration.ManagedIdentityAccountId = ReadOptional(getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.ManagedIdentityAccountId));
        configuration.GraphTenantId = ReadOptional(getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.GraphTenantId));
        configuration.DefaultUsageLocation = NormalizeUsageLocation(getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.DefaultUsageLocation));

        var authModeRaw = getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.AuthenticationMode);
        if (!string.IsNullOrWhiteSpace(authModeRaw))
        {
            configuration.AuthenticationMode = ParseAuthenticationModeOrThrow(authModeRaw);
        }

        var scopesRaw = getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.GraphScopes);
        if (!string.IsNullOrWhiteSpace(scopesRaw))
        {
            configuration.GraphScopes = scopesRaw
                .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .ToList();
        }

        var writeScopesRaw = getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.GraphLicenseWriteScopes);
        if (!string.IsNullOrWhiteSpace(writeScopesRaw))
        {
            configuration.GraphLicenseWriteScopes = writeScopesRaw
                .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .ToList();
        }

        var enableGraphRaw = getEnvironmentVariable(ExchangeConfigurationEnvironmentVariables.EnableGraphAfterExchangeConnect);
        if (!string.IsNullOrWhiteSpace(enableGraphRaw))
        {
            configuration.EnableGraphAfterExchangeConnect =
                string.Equals(enableGraphRaw, "1", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(enableGraphRaw, "true", StringComparison.OrdinalIgnoreCase);
        }

        configuration.GraphScopes = configuration.NormalizeGraphScopes().ToList();
        if (configuration.GraphLicenseWriteScopes is not null)
        {
            configuration.GraphLicenseWriteScopes = configuration.NormalizeGraphLicenseWriteScopes().ToList();
        }

        configuration.DefaultUsageLocation = NormalizeUsageLocation(configuration.DefaultUsageLocation);

        return configuration;
    }

    public static ExchangeAuthenticationMode ParseAuthenticationModeOrThrow(string value)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(value);

        if (Enum.TryParse<ExchangeAuthenticationMode>(value.Trim(), ignoreCase: true, out var authenticationMode))
        {
            return authenticationMode;
        }

        throw new InvalidOperationException(
            $"AuthenticationMode must be one of: {string.Join(", ", Enum.GetNames<ExchangeAuthenticationMode>())}.");
    }

    private static string ReadOrDefault(string? value, string defaultValue)
        => string.IsNullOrWhiteSpace(value) ? defaultValue : value.Trim();

    private static string? ReadOptional(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    private static string? NormalizeOptional(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    private static string? ValidateExchangeOrganizationValue(string? value)
        => ValidateTenantIdentifier(value, "ExchangeOrganization");

    private static string? ValidateGraphTenantIdValue(string? value)
        => ValidateTenantIdentifier(value, "GraphTenantId");

    public static string? NormalizeUsageLocation(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : value.Trim().ToUpperInvariant();

    public static bool IsValidUsageLocation(string? value)
        => ValidateUsageLocationValue(value, "UsageLocation") is null;

    private static string? ValidateUsageLocationValue(string? value, string fieldName)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var normalizedValue = NormalizeUsageLocation(value);
        return normalizedValue is { Length: 2 } && normalizedValue.All(char.IsLetter)
            ? null
            : $"{fieldName} must be a valid two-letter country/region code.";
    }

    private static string? ValidateTenantIdentifier(string? value, string fieldName)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var normalizedValue = value.Trim();
        if (Guid.TryParse(normalizedValue, out _))
        {
            return null;
        }

        if (LooksLikeDnsName(normalizedValue))
        {
            return null;
        }

        return $"{fieldName} must be a tenant domain like 'contoso.onmicrosoft.com' or a tenant GUID.";
    }

    private static bool LooksLikeDnsName(string value)
    {
        if (value.Length > 253 ||
            value.StartsWith(".", StringComparison.Ordinal) ||
            value.EndsWith(".", StringComparison.Ordinal) ||
            !value.Contains(".", StringComparison.Ordinal))
        {
            return false;
        }

        var labels = value.Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (labels.Length < 2)
        {
            return false;
        }

        foreach (var label in labels)
        {
            if (label.Length is 0 or > 63 ||
                label.StartsWith("-", StringComparison.Ordinal) ||
                label.EndsWith("-", StringComparison.Ordinal))
            {
                return false;
            }

            foreach (var character in label)
            {
                if (!(char.IsLetterOrDigit(character) || character == '-'))
                {
                    return false;
                }
            }
        }

        return true;
    }
}

