using ExchangeAdmin.Contracts;

namespace ExchangeAdmin.Worker.PowerShell;

internal static class GraphCommandBuilder
{
    internal static string BuildConnectGraphCommand(
        ExchangeOnlineConfiguration configuration,
        IEnumerable<string>? delegatedScopes = null)
    {
        ArgumentNullException.ThrowIfNull(configuration);
        ExchangeOnlineConfiguration.ThrowIfInvalidGraphTenantId(configuration.GraphTenantId);

        return configuration.AuthenticationMode switch
        {
            ExchangeAuthenticationMode.Interactive => BuildInteractiveCommand(configuration, delegatedScopes),
            ExchangeAuthenticationMode.DeviceCode => BuildDeviceCodeCommand(configuration, delegatedScopes),
            ExchangeAuthenticationMode.AppCertificate => BuildAppCertificateCommand(configuration),
            ExchangeAuthenticationMode.ManagedIdentity => BuildManagedIdentityCommand(configuration),
            _ => throw new InvalidOperationException($"Unsupported Graph authentication mode '{configuration.AuthenticationMode}'.")
        };
    }

    private static string BuildInteractiveCommand(
        ExchangeOnlineConfiguration configuration,
        IEnumerable<string>? delegatedScopes)
    {
        var commandParts = new List<string>
        {
            "Connect-MgGraph",
            BuildScopesArgument(configuration.NormalizeGraphScopes(delegatedScopes))
        };

        if (!string.IsNullOrWhiteSpace(configuration.GraphTenantId))
        {
            commandParts.Add($"-TenantId '{EscapePs(configuration.GraphTenantId)}'");
        }

        commandParts.Add("-ContextScope Process");
        commandParts.Add("-NoWelcome");
        return string.Join(" ", commandParts);
    }

    private static string BuildDeviceCodeCommand(
        ExchangeOnlineConfiguration configuration,
        IEnumerable<string>? delegatedScopes)
    {
        var commandParts = new List<string>
        {
            "Connect-MgGraph",
            BuildScopesArgument(configuration.NormalizeGraphScopes(delegatedScopes))
        };

        if (!string.IsNullOrWhiteSpace(configuration.GraphTenantId))
        {
            commandParts.Add($"-TenantId '{EscapePs(configuration.GraphTenantId)}'");
        }

        commandParts.Add("-ContextScope Process");
        commandParts.Add("-UseDeviceCode");
        commandParts.Add("-NoWelcome");
        return string.Join(" ", commandParts);
    }

    private static string BuildAppCertificateCommand(ExchangeOnlineConfiguration configuration)
    {
        if (string.IsNullOrWhiteSpace(configuration.ApplicationId) || string.IsNullOrWhiteSpace(configuration.GraphTenantId))
        {
            throw new InvalidOperationException("Graph app certificate authentication requires ApplicationId and GraphTenantId.");
        }

        var commandParts = new List<string>
        {
            "Connect-MgGraph",
            $"-ClientId '{EscapePs(configuration.ApplicationId)}'",
            $"-TenantId '{EscapePs(configuration.GraphTenantId)}'"
        };

        if (!string.IsNullOrWhiteSpace(configuration.CertificateThumbprint))
        {
            commandParts.Add($"-CertificateThumbprint '{EscapePs(configuration.CertificateThumbprint)}'");
        }
        else if (!string.IsNullOrWhiteSpace(configuration.CertificateSubjectName))
        {
            commandParts.Add($"-CertificateSubjectName '{EscapePs(configuration.CertificateSubjectName)}'");
        }
        else
        {
            throw new InvalidOperationException("Graph app certificate authentication requires CertificateThumbprint or CertificateSubjectName.");
        }

        commandParts.Add("-ContextScope Process");
        commandParts.Add("-NoWelcome");
        return string.Join(" ", commandParts);
    }

    private static string BuildManagedIdentityCommand(ExchangeOnlineConfiguration configuration)
    {
        var commandParts = new List<string> { "Connect-MgGraph", "-Identity" };

        if (!string.IsNullOrWhiteSpace(configuration.ManagedIdentityAccountId))
        {
            commandParts.Add($"-ClientId '{EscapePs(configuration.ManagedIdentityAccountId)}'");
        }

        commandParts.Add("-ContextScope Process");
        commandParts.Add("-NoWelcome");
        return string.Join(" ", commandParts);
    }

    private static string BuildScopesArgument(IEnumerable<string> scopes)
    {
        var values = scopes.Select(scope => $"'{EscapePs(scope)}'");
        return $"-Scopes @({string.Join(", ", values)})";
    }

    private static string EscapePs(string value) => value.Replace("'", "''", StringComparison.Ordinal);
}
