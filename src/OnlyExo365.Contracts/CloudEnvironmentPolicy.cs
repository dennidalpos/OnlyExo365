using System.Text.Json;

namespace OnlyExo365.Contracts;

public sealed record CloudEnvironmentRoute(string GraphEnvironment, string ComplianceUri, string Authority);

public static class CloudEnvironmentPolicy
{
    private static readonly IReadOnlyDictionary<string, CloudEnvironmentRoute> Routes = Load();
    private static IReadOnlyDictionary<string, CloudEnvironmentRoute> Load()
    {
        using var stream = typeof(CloudEnvironmentPolicy).Assembly.GetManifestResourceStream("OnlyExo365.Contracts.Data.CloudEnvironments.json")!;
        return new Dictionary<string, CloudEnvironmentRoute>(
            JsonSerializer.Deserialize<Dictionary<string, CloudEnvironmentRoute>>(stream)!, StringComparer.OrdinalIgnoreCase);
    }

    public static CloudEnvironmentRoute Resolve(string? exchangeEnvironment)
    {
        var key = string.IsNullOrWhiteSpace(exchangeEnvironment) ? "O365Default" : exchangeEnvironment.Trim();
        return Routes.TryGetValue(key, out var route) ? route
            : throw new InvalidOperationException($"Unsupported Microsoft 365 cloud: '{key}'.");
    }

    public static string BuildComplianceRoutingArguments(ExchangeOnlineConfiguration configuration)
    {
        var route = Resolve(configuration.ExchangeEnvironmentName);
        var tenant = string.IsNullOrWhiteSpace(configuration.DelegatedOrganization) ? "organizations" : Uri.EscapeDataString(configuration.DelegatedOrganization);
        return $"-ConnectionUri '{route.ComplianceUri}' -AzureADAuthorizationEndpointUri '{route.Authority}/{tenant}'";
    }

    public static void ValidateGraphAuthentication(ExchangeOnlineConfiguration configuration)
    {
        if (Resolve(configuration.ExchangeEnvironmentName).GraphEnvironment == "China" &&
            configuration.AuthenticationMode is ExchangeAuthenticationMode.Interactive or ExchangeAuthenticationMode.DeviceCode &&
            string.IsNullOrWhiteSpace(configuration.ApplicationId))
            throw new InvalidOperationException("Delegated Graph authentication in China requires a locally registered ApplicationId.");
    }
}
