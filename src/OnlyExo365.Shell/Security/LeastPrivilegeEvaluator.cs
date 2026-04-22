using OnlyExo365.Contracts;
using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Shell.Security;

public sealed class LeastPrivilegeEvaluator
{
    private readonly ExchangeOnlineConfiguration _configuration;

    public LeastPrivilegeEvaluator(ExchangeOnlineConfiguration configuration)
    {
        _configuration = configuration?.Clone() ?? ExchangeOnlineConfiguration.CreateDefault();
    }

    public IReadOnlyList<LeastPrivilegeFeatureEvaluation> EvaluateAll(CapabilityMapDto? capabilities)
        => LeastPrivilegeCatalog.All.Select(feature => Evaluate(feature, capabilities)).ToList();

    public LeastPrivilegeFeatureEvaluation Evaluate(string featureId, CapabilityMapDto? capabilities)
    {
        var feature = LeastPrivilegeCatalog.All.FirstOrDefault(item =>
            string.Equals(item.FeatureId, featureId, StringComparison.Ordinal));

        if (feature == null)
        {
            throw new InvalidOperationException($"Unknown least-privilege feature id '{featureId}'.");
        }

        return Evaluate(feature, capabilities);
    }

    private LeastPrivilegeFeatureEvaluation Evaluate(
        LeastPrivilegeFeatureDefinition feature,
        CapabilityMapDto? capabilities)
    {
        if (capabilities == null)
        {
            return new LeastPrivilegeFeatureEvaluation(
                feature,
                LeastPrivilegeFeatureStatus.PendingSession,
                Array.Empty<string>(),
                "Connect to Exchange and complete capability detection to validate cmdlets for this feature.");
        }

        var missingRequirements = new List<string>();

        if (feature.AllowedAuthenticationModes.Count > 0 &&
            !feature.AllowedAuthenticationModes.Contains(_configuration.AuthenticationMode))
        {
            missingRequirements.Add($"Authentication mode '{_configuration.AuthenticationMode}' is not approved for this feature.");
        }

        var configuredScopes = _configuration.NormalizeGraphScopes()
            .Concat(_configuration.NormalizeGraphLicenseWriteScopes())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        foreach (var requiredScope in feature.RequiredGraphScopes)
        {
            if (!configuredScopes.Contains(requiredScope, StringComparer.OrdinalIgnoreCase))
            {
                missingRequirements.Add($"Missing Graph scope '{requiredScope}'.");
            }
        }

        if (feature.RequiredGraphScopes.Count > 0 &&
            _configuration.AuthenticationMode == ExchangeAuthenticationMode.AppCertificate &&
            string.IsNullOrWhiteSpace(_configuration.GraphTenantId))
        {
            missingRequirements.Add("GraphTenantId is required for AppCertificate Graph access.");
        }

        foreach (var cmdlet in feature.RequiredCmdletsAll)
        {
            if (!IsCmdletAvailable(capabilities, cmdlet))
            {
                missingRequirements.Add(BuildCmdletMessage(capabilities, cmdlet));
            }
        }

        if (feature.RequiredCmdletsAny.Count > 0 &&
            !feature.RequiredCmdletsAny.Any(cmdlet => IsCmdletAvailable(capabilities, cmdlet)))
        {
            missingRequirements.Add(
                $"None of the alternative cmdlets are available: {string.Join(", ", feature.RequiredCmdletsAny)}.");
        }

        if (missingRequirements.Count == 0)
        {
            return new LeastPrivilegeFeatureEvaluation(
                feature,
                LeastPrivilegeFeatureStatus.Available,
                Array.Empty<string>(),
                "Current configuration satisfies the validated least-privilege checks.");
        }

        if (feature.RequiresAdditionalSessionValidation &&
            missingRequirements.All(item => item.Contains("cmdlet", StringComparison.OrdinalIgnoreCase) ||
                                            item.Contains("alternative cmdlets", StringComparison.OrdinalIgnoreCase)))
        {
            return new LeastPrivilegeFeatureEvaluation(
                feature,
                LeastPrivilegeFeatureStatus.NeedsAdditionalSession,
                missingRequirements,
                "This feature uses a secondary session. Validate again after the module connects to its dedicated endpoint.");
        }

        return new LeastPrivilegeFeatureEvaluation(
            feature,
            LeastPrivilegeFeatureStatus.Blocked,
            missingRequirements,
            "Current configuration does not satisfy the validated least-privilege checks.");
    }

    private static bool IsCmdletAvailable(CapabilityMapDto capabilities, string cmdletName)
        => capabilities.Cmdlets.TryGetValue(cmdletName, out var cmdlet) && cmdlet.IsAvailable;

    private static string BuildCmdletMessage(CapabilityMapDto capabilities, string cmdletName)
    {
        if (!capabilities.Cmdlets.TryGetValue(cmdletName, out var cmdlet))
        {
            return $"Required cmdlet '{cmdletName}' was not detected.";
        }

        if (string.IsNullOrWhiteSpace(cmdlet.UnavailableReason))
        {
            return $"Required cmdlet '{cmdletName}' is unavailable.";
        }

        return $"Required cmdlet '{cmdletName}' is unavailable: {cmdlet.UnavailableReason}";
    }
}

