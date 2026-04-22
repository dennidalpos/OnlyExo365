using System.Reflection;
using System.Text.Json;

namespace OnlyExo365.Worker.PowerShell;

internal static class PowerShellModuleBootstrapPolicy
{
    private const string ResourceName = "OnlyExo365.Worker.Data.PowerShellModuleBootstrapPolicy.json";
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };
    private static readonly Lazy<PolicyDefinition> Policy = new(LoadPolicy);

    public static string RepositoryName => Policy.Value.RepositoryName;

    public static string RepositorySourceLocation => Policy.Value.RepositorySourceLocation;

    public static PowerShellModuleDefinition? Resolve(string? requestedModuleName)
    {
        if (string.IsNullOrWhiteSpace(requestedModuleName))
        {
            return null;
        }

        var normalizedName = requestedModuleName.Trim();
        var definition = Policy.Value.Modules.FirstOrDefault(module =>
            string.Equals(module.RequestName, normalizedName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(module.ModuleName, normalizedName, StringComparison.OrdinalIgnoreCase) ||
            module.Aliases.Any(alias => string.Equals(alias, normalizedName, StringComparison.OrdinalIgnoreCase)));

        return definition;
    }

    private static PolicyDefinition LoadPolicy()
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName)
            ?? throw new InvalidOperationException($"Embedded resource not found: {ResourceName}");
        using var reader = new StreamReader(stream);
        var json = reader.ReadToEnd();

        var policy = JsonSerializer.Deserialize<PolicyDefinition>(json, SerializerOptions)
            ?? throw new InvalidOperationException("Unable to deserialize PowerShell module bootstrap policy.");

        if (string.IsNullOrWhiteSpace(policy.RepositoryName) ||
            string.IsNullOrWhiteSpace(policy.RepositorySourceLocation) ||
            policy.Modules.Count == 0)
        {
            throw new InvalidOperationException("PowerShell module bootstrap policy is incomplete.");
        }

        foreach (var module in policy.Modules)
        {
            if (string.IsNullOrWhiteSpace(module.RequestName) ||
                string.IsNullOrWhiteSpace(module.ModuleName) ||
                string.IsNullOrWhiteSpace(module.RequiredVersion))
            {
                throw new InvalidOperationException("PowerShell module bootstrap policy contains an incomplete module entry.");
            }
        }

        return policy;
    }

    internal sealed class PolicyDefinition
    {
        public string RepositoryName { get; set; } = string.Empty;

        public string RepositorySourceLocation { get; set; } = string.Empty;

        public List<PowerShellModuleDefinition> Modules { get; set; } = [];
    }
}

internal sealed class PowerShellModuleDefinition
{
    public string RequestName { get; set; } = string.Empty;

    public string ModuleName { get; set; } = string.Empty;

    public string RequiredVersion { get; set; } = string.Empty;

    public List<string> Aliases { get; set; } = [];

    public List<string> RequiredModules { get; set; } = [];

    public IReadOnlyList<string> GetRequiredModules()
    {
        if (RequiredModules.Count == 0)
        {
            return [ModuleName];
        }

        return RequiredModules
            .Where(static module => !string.IsNullOrWhiteSpace(module))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public string BuildManualInstructions(string repositoryName)
    {
        var modules = GetRequiredModules();
        if (modules.Count == 1)
        {
            return
                $"If available, run `Install-PSResource {modules[0]} -Version {RequiredVersion} -Repository {repositoryName} -Scope CurrentUser -TrustRepository`. " +
                $"Alternatively, download `https://www.powershellgallery.com/api/v2/package/{modules[0]}/{RequiredVersion}` " +
                $"and extract the content to `$HOME\\Documents\\PowerShell\\Modules\\{modules[0]}\\{RequiredVersion}`.";
        }

        var installCommands = string.Join(" ", modules.Select(module =>
            $"`Install-PSResource {module} -Version {RequiredVersion} -Repository {repositoryName} -Scope CurrentUser -TrustRepository`"));
        var packageUris = string.Join(" ", modules.Select(module =>
            $"`https://www.powershellgallery.com/api/v2/package/{module}/{RequiredVersion}`"));
        return
            $"Install all approved Graph modules at version {RequiredVersion}: {installCommands}. " +
            $"Alternatively, download the packages {packageUris} and extract each package content to `$HOME\\Documents\\PowerShell\\Modules\\<ModuleName>\\{RequiredVersion}`.";
    }
}

