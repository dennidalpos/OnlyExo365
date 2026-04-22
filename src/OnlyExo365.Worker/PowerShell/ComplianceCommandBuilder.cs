using System.Text;
using OnlyExo365.Contracts;
using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

internal static class ComplianceCommandBuilder
{
    internal static string BuildConnectComplianceCommand(ExchangeOnlineConfiguration configuration)
        => BuildConnectComplianceCommand(configuration, enableSearchOnlySession: false);

    internal static string BuildConnectComplianceSearchOnlyCommand(ExchangeOnlineConfiguration configuration)
        => BuildConnectComplianceCommand(configuration, enableSearchOnlySession: true);

    private static string BuildConnectComplianceCommand(ExchangeOnlineConfiguration configuration, bool enableSearchOnlySession)
    {
        ArgumentNullException.ThrowIfNull(configuration);
        ExchangeOnlineConfiguration.ThrowIfInvalidExchangeOrganization(configuration.ExchangeOrganization);

        return configuration.AuthenticationMode switch
        {
            ExchangeAuthenticationMode.Interactive => BuildDelegatedCommand(configuration, includeDeviceSwitch: false, enableSearchOnlySession),
            ExchangeAuthenticationMode.DeviceCode => BuildDelegatedCommand(configuration, includeDeviceSwitch: true, enableSearchOnlySession),
            ExchangeAuthenticationMode.AppCertificate => BuildAppCertificateCommand(configuration, enableSearchOnlySession),
            ExchangeAuthenticationMode.ManagedIdentity => BuildManagedIdentityCommand(configuration, enableSearchOnlySession),
            _ => throw new InvalidOperationException($"Unsupported compliance authentication mode '{configuration.AuthenticationMode}'.")
        };
    }

    internal static string BuildCreateComplianceSearchScript(CreateComplianceSearchRequest request)
    {
        var escapedName = EscapePs(request.Name);
        var escapedCase = EscapePs(request.CaseName);
        var escapedQuery = EscapePs(request.ContentMatchQuery);
        var exchangeLocations = ToPsArrayLiteral(request.ExchangeLocations);

        var script = new StringBuilder();
        script.AppendLine("$params = @{");
        script.AppendLine($"    Name = '{escapedName}'");
        script.AppendLine($"    ExchangeLocation = {exchangeLocations}");
        script.AppendLine("}");
        script.AppendLine($"if ('{escapedCase}' -ne '') {{ $params['Case'] = '{escapedCase}' }}");
        script.AppendLine($"if ('{escapedQuery}' -ne '') {{ $params['ContentMatchQuery'] = '{escapedQuery}' }}");
        script.AppendLine("New-ComplianceSearch @params -ErrorAction Stop | Out-Null");
        script.AppendLine("Get-ComplianceSearch -Identity $params['Name'] -ErrorAction Stop | Select-Object -First 1");
        return script.ToString();
    }

    internal static string BuildStartComplianceSearchScript(string searchName)
    {
        var escapedSearchName = EscapePs(searchName);

        return $@"
Start-ComplianceSearch -Identity '{escapedSearchName}' -ErrorAction Stop | Out-Null
Get-ComplianceSearch -Identity '{escapedSearchName}' -ErrorAction Stop | Select-Object -First 1";
    }

    internal static string BuildRemoveComplianceSearchScript(string searchName)
    {
        var escapedSearchName = EscapePs(searchName);

        return $"Remove-ComplianceSearch -Identity '{escapedSearchName}' -Confirm:$false -ErrorAction Stop";
    }

    internal static string BuildInvokeComplianceActionScript(
        InvokeComplianceActionRequest request,
        IReadOnlyList<string> exchangeLocations,
        string? contentMatchQuery)
    {
        var escapedSearchName = EscapePs(request.SearchName);
        var escapedActionType = EscapePs(request.ActionType);
        var escapedPurgeType = EscapePs(request.PurgeType);
        var escapedCaseName = EscapePs(request.CaseName);
        var escapedHoldName = EscapePs(request.HoldName);
        var escapedQuery = EscapePs(contentMatchQuery);
        var escapedRuleName = EscapePs(string.IsNullOrWhiteSpace(request.HoldName) ? $"{request.SearchName} Rule" : $"{request.HoldName} Rule");
        var exchangeLocationArray = ToPsArrayLiteral(exchangeLocations);

        var script = new StringBuilder();
        script.AppendLine("$result = $null");
        script.AppendLine($"switch ('{escapedActionType}') {{");
        script.AppendLine("    'Purge' {");
        script.AppendLine($"        $result = New-ComplianceSearchAction -SearchName '{escapedSearchName}' -Purge -PurgeType '{escapedPurgeType}' -Confirm:$false -ErrorAction Stop");
        script.AppendLine("        break");
        script.AppendLine("    }");
        script.AppendLine("    'Hold' {");
        script.AppendLine($"        $holdName = '{escapedHoldName}'");
        script.AppendLine($"        $caseName = '{escapedCaseName}'");
        script.AppendLine($"        $locations = {exchangeLocationArray}");
        script.AppendLine("        if ([string]::IsNullOrWhiteSpace($holdName)) { throw 'HoldName is required for hold actions.' }");
        script.AppendLine("        if ([string]::IsNullOrWhiteSpace($caseName)) { throw 'CaseName is required for hold actions.' }");
        script.AppendLine("        if ($locations.Count -eq 0) { throw 'At least one Exchange location is required to create a case hold.' }");
        script.AppendLine("        New-CaseHoldPolicy -Name $holdName -Case $caseName -ExchangeLocation $locations -ErrorAction Stop | Out-Null");
        script.AppendLine($"        if ('{escapedQuery}' -ne '') {{");
        script.AppendLine($"            New-CaseHoldRule -Name '{escapedRuleName}' -Policy $holdName -ContentMatchQuery '{escapedQuery}' -ErrorAction Stop | Out-Null");
        script.AppendLine("        }");
        script.AppendLine("        $result = Get-CaseHoldPolicy -Identity $holdName -Case $caseName -DistributionDetail -ErrorAction Stop | Select-Object -First 1");
        script.AppendLine("        break");
        script.AppendLine("    }");
        script.AppendLine("    default {");
        script.AppendLine("        throw 'Unsupported compliance action type.'");
        script.AppendLine("    }");
        script.AppendLine("}");
        script.AppendLine("$result");
        return script.ToString();
    }

    private static string BuildDelegatedCommand(
        ExchangeOnlineConfiguration configuration,
        bool includeDeviceSwitch,
        bool enableSearchOnlySession)
    {
        var commandParts = BuildBaseCommand(configuration, enableSearchOnlySession);

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

    private static string BuildAppCertificateCommand(ExchangeOnlineConfiguration configuration, bool enableSearchOnlySession)
    {
        if (string.IsNullOrWhiteSpace(configuration.ApplicationId) || string.IsNullOrWhiteSpace(configuration.ExchangeOrganization))
        {
            throw new InvalidOperationException("Compliance app certificate authentication requires ApplicationId and ExchangeOrganization.");
        }

        var commandParts = BuildBaseCommand(configuration, enableSearchOnlySession);
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
            throw new InvalidOperationException("Compliance app certificate authentication requires CertificateThumbprint or CertificateSubjectName.");
        }

        return string.Join(" ", commandParts);
    }

    private static string BuildManagedIdentityCommand(ExchangeOnlineConfiguration configuration, bool enableSearchOnlySession)
    {
        if (string.IsNullOrWhiteSpace(configuration.ExchangeOrganization))
        {
            throw new InvalidOperationException("Compliance managed identity authentication requires ExchangeOrganization.");
        }

        var commandParts = BuildBaseCommand(configuration, enableSearchOnlySession);
        commandParts.Add("-ManagedIdentity");
        commandParts.Add($"-Organization '{EscapePs(configuration.ExchangeOrganization)}'");

        if (!string.IsNullOrWhiteSpace(configuration.ManagedIdentityAccountId))
        {
            commandParts.Add($"-ManagedIdentityAccountId '{EscapePs(configuration.ManagedIdentityAccountId)}'");
        }

        return string.Join(" ", commandParts);
    }

    private static List<string> BuildBaseCommand(ExchangeOnlineConfiguration configuration, bool enableSearchOnlySession)
    {
        var commandParts = new List<string> { "Connect-IPPSSession", "-WarningAction", "SilentlyContinue" };

        if (enableSearchOnlySession)
        {
            commandParts.Add("-EnableSearchOnlySession");
        }

        if (!string.IsNullOrWhiteSpace(configuration.ExchangeOrganization))
        {
            commandParts.Add($"-Organization '{EscapePs(configuration.ExchangeOrganization)}'");
        }

        if (!string.IsNullOrWhiteSpace(configuration.DelegatedOrganization))
        {
            commandParts.Add($"-DelegatedOrganization '{EscapePs(configuration.DelegatedOrganization)}'");
        }

        return commandParts;
    }

    private static string EscapePs(string? value)
        => (value ?? string.Empty).Replace("'", "''", StringComparison.Ordinal);

    private static string ToPsArrayLiteral(IEnumerable<string>? values)
    {
        if (values == null)
        {
            return "@()";
        }

        var normalized = values
            .Where(static value => !string.IsNullOrWhiteSpace(value))
            .Select(static value => $"'{EscapePs(value.Trim())}'")
            .ToArray();

        return normalized.Length == 0
            ? "@()"
            : "@(" + string.Join(", ", normalized) + ")";
    }
}

