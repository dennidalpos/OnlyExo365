using OnlyExo365.Contracts;

namespace OnlyExo365.Worker.PowerShell;

internal static class ExchangeCommandBuilder
{
    private static readonly HashSet<string> SupportedExchangeEnvironments = new(StringComparer.OrdinalIgnoreCase)
    {
        "O365Default",
        "O365GermanyCloud",
        "O365USGovGCCHigh",
        "O365USGovDoD",
        "O365China"
    };

    internal static string BuildConnectExchangeCommand(ExchangeOnlineConfiguration configuration)
    {
        ArgumentNullException.ThrowIfNull(configuration);
        ExchangeOnlineConfiguration.ThrowIfInvalidExchangeOrganization(configuration.ExchangeOrganization);

        return configuration.AuthenticationMode switch
        {
            ExchangeAuthenticationMode.Interactive => BuildDelegatedCommand(configuration, includeDeviceSwitch: false),
            ExchangeAuthenticationMode.DeviceCode => BuildDelegatedCommand(configuration, includeDeviceSwitch: true),
            ExchangeAuthenticationMode.AppCertificate => BuildAppCertificateCommand(configuration),
            ExchangeAuthenticationMode.ManagedIdentity => BuildManagedIdentityCommand(configuration),
            _ => throw new InvalidOperationException($"Unsupported Exchange authentication mode '{configuration.AuthenticationMode}'.")
        };
    }

    private static string BuildDelegatedCommand(ExchangeOnlineConfiguration configuration, bool includeDeviceSwitch)
    {
        var commandParts = BuildBaseCommand(configuration);

        if (includeDeviceSwitch)
        {
            commandParts.Add("-Device");
        }

        if (!string.IsNullOrWhiteSpace(configuration.UserPrincipalNameHint))
        {
            commandParts.Add($"-UserPrincipalName '{EscapePs(configuration.UserPrincipalNameHint)}'");
        }

        return string.Join(" ", commandParts);
    }

    private static string BuildAppCertificateCommand(ExchangeOnlineConfiguration configuration)
    {
        if (string.IsNullOrWhiteSpace(configuration.ApplicationId) || string.IsNullOrWhiteSpace(configuration.ExchangeOrganization))
        {
            throw new InvalidOperationException("Exchange app certificate authentication requires ApplicationId and ExchangeOrganization.");
        }

        var commandParts = BuildBaseCommand(configuration);
        commandParts.Add($"-AppId '{EscapePs(configuration.ApplicationId)}'");
        commandParts.Add($"-Organization '{EscapePs(configuration.ExchangeOrganization)}'");

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
            throw new InvalidOperationException("Exchange app certificate authentication requires CertificateThumbprint or CertificateSubjectName.");
        }

        return string.Join(" ", commandParts);
    }

    private static string BuildManagedIdentityCommand(ExchangeOnlineConfiguration configuration)
    {
        if (string.IsNullOrWhiteSpace(configuration.ExchangeOrganization))
        {
            throw new InvalidOperationException("Exchange managed identity authentication requires ExchangeOrganization.");
        }

        var commandParts = BuildBaseCommand(configuration);
        commandParts.Add("-ManagedIdentity");
        commandParts.Add($"-Organization '{EscapePs(configuration.ExchangeOrganization)}'");

        if (!string.IsNullOrWhiteSpace(configuration.ManagedIdentityAccountId))
        {
            commandParts.Add($"-ManagedIdentityAccountId '{EscapePs(configuration.ManagedIdentityAccountId)}'");
        }

        return string.Join(" ", commandParts);
    }

    private static List<string> BuildBaseCommand(ExchangeOnlineConfiguration configuration)
    {
        var commandParts = new List<string> { "Connect-ExchangeOnline", "-ShowBanner:$false" };
        var exchangeEnvironment = configuration.ExchangeEnvironmentName;

        if (!string.IsNullOrWhiteSpace(exchangeEnvironment) &&
            SupportedExchangeEnvironments.Contains(exchangeEnvironment))
        {
            commandParts.Add($"-ExchangeEnvironmentName '{EscapePs(exchangeEnvironment)}'");
        }

        if (!string.IsNullOrWhiteSpace(configuration.ExchangeOrganization) &&
            configuration.AuthenticationMode is ExchangeAuthenticationMode.Interactive or ExchangeAuthenticationMode.DeviceCode)
        {
            commandParts.Add($"-Organization '{EscapePs(configuration.ExchangeOrganization)}'");
        }

        if (!string.IsNullOrWhiteSpace(configuration.DelegatedOrganization))
        {
            commandParts.Add($"-DelegatedOrganization '{EscapePs(configuration.DelegatedOrganization)}'");
        }

        return commandParts;
    }

    private static string EscapePs(string value) => value.Replace("'", "''", StringComparison.Ordinal);
}

