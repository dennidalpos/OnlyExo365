using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

internal static class MigrationCommandBuilder
{
    public static (string Script, Dictionary<string, object>? Parameters) BuildUpsertMigrationEndpointCommand(UpsertMigrationEndpointRequest request)
    {
        var endpointType = NormalizeEndpointType(request.EndpointType);
        var isUpdate = !string.IsNullOrWhiteSpace(request.Identity);
        var parameters = new Dictionary<string, object>();
        var lines = new List<string>();

        if (RequiresCredential(endpointType) &&
            !string.IsNullOrWhiteSpace(request.Username) &&
            !string.IsNullOrWhiteSpace(request.Password))
        {
            parameters["MigrationUserName"] = request.Username.Trim();
            parameters["MigrationPassword"] = request.Password;
            lines.Add("$securePassword = ConvertTo-SecureString $MigrationPassword -AsPlainText -Force");
            lines.Add("$credential = [pscredential]::new($MigrationUserName, $securePassword)");
        }

        lines.Add("$params = @{}");

        if (isUpdate)
        {
            lines.Add($"$params['Identity'] = '{EscapePs(request.Identity)}'");
        }
        else
        {
            lines.Add($"$params['Name'] = '{EscapePs(request.Name)}'");
            lines.Add($"$params['{endpointType}'] = $true");
        }

        AddStringParam(lines, "RemoteServer", request.RemoteServer);
        AddStringParam(lines, "RpcProxyServer", request.RpcProxyServer);
        AddStringParam(lines, "ExchangeServer", request.ExchangeServer);
        AddStringParam(lines, "EmailAddress", request.EmailAddress);
        AddStringParam(lines, "RemoteTenant", isUpdate ? null : request.RemoteTenant);
        AddNullableIntParam(lines, "Port", request.Port);
        AddStringParam(lines, "Security", endpointType == "IMAP" ? request.Security : null);
        AddStringParam(lines, "Authentication", endpointType == "ExchangeOutlookAnywhere" ? request.Authentication : null);
        AddNullableIntParam(lines, "MaxConcurrentMigrations", request.MaxConcurrentMigrations);
        AddNullableIntParam(lines, "MaxConcurrentIncrementalSyncs", request.MaxConcurrentIncrementalSyncs);

        if (request.SkipVerification)
        {
            lines.Add("$params['SkipVerification'] = $true");
        }

        if (request.AcceptUntrustedCertificates)
        {
            lines.Add("$params['AcceptUntrustedCertificates'] = $true");
        }

        if (parameters.ContainsKey("MigrationUserName"))
        {
            lines.Add("$params['Credentials'] = $credential");
        }

        lines.Add(isUpdate
            ? "Set-MigrationEndpoint @params -ErrorAction Stop | Out-Null"
            : "New-MigrationEndpoint @params -ErrorAction Stop | Out-Null");
        lines.Add("Write-Output 'OK'");

        return (string.Join(Environment.NewLine, lines), parameters.Count == 0 ? null : parameters);
    }

    public static (string Script, Dictionary<string, object>? Parameters) BuildTestMigrationEndpointCommand(TestMigrationEndpointRequest request)
    {
        if (request.UseExistingEndpoint && !string.IsNullOrWhiteSpace(request.Identity))
        {
            var script = $@"
$result = Test-MigrationServerAvailability -Endpoint '{EscapePs(request.Identity)}' -ErrorAction Stop | Select-Object -First 1
[PSCustomObject]@{{
    Summary = 'Migration endpoint test completed successfully.'
    Details = if ($result) {{ $result | Out-String }} else {{ $null }}
}}";

            return (script, null);
        }

        var endpointType = NormalizeEndpointType(request.EndpointType);
        var parameters = new Dictionary<string, object>();
        var lines = new List<string>();

        if (RequiresCredential(endpointType))
        {
            parameters["MigrationUserName"] = request.Username?.Trim() ?? string.Empty;
            parameters["MigrationPassword"] = request.Password ?? string.Empty;
            lines.Add("$securePassword = ConvertTo-SecureString $MigrationPassword -AsPlainText -Force");
            lines.Add("$credential = [pscredential]::new($MigrationUserName, $securePassword)");
        }

        lines.Add("$params = @{}");
        lines.Add($"$params['{endpointType}'] = $true");
        AddStringParam(lines, "RemoteServer", request.RemoteServer);
        AddStringParam(lines, "RpcProxyServer", request.RpcProxyServer);
        AddStringParam(lines, "ExchangeServer", request.ExchangeServer);
        AddStringParam(lines, "EmailAddress", request.EmailAddress);
        AddNullableIntParam(lines, "Port", request.Port);
        AddStringParam(lines, "Security", endpointType == "IMAP" ? request.Security : null);
        AddStringParam(lines, "Authentication", endpointType == "ExchangeOutlookAnywhere" ? request.Authentication : null);

        if (request.SkipVerification)
        {
            lines.Add("$params['SkipVerification'] = $true");
        }

        if (request.AcceptUntrustedCertificates)
        {
            lines.Add("$params['AcceptUntrustedCertificates'] = $true");
        }

        if (parameters.ContainsKey("MigrationUserName"))
        {
            lines.Add("$params['Credentials'] = $credential");
        }

        lines.Add("$result = Test-MigrationServerAvailability @params -ErrorAction Stop | Select-Object -First 1");
        lines.Add("[PSCustomObject]@{");
        lines.Add("    Summary = 'Migration endpoint test completed successfully.'");
        lines.Add("    Details = if ($result) { $result | Out-String } else { $null }");
        lines.Add("}");

        return (string.Join(Environment.NewLine, lines), parameters);
    }

    public static (string Script, Dictionary<string, object>? Parameters) BuildCreateMigrationBatchCommand(CreateMigrationBatchRequest request)
    {
        var batchType = NormalizeBatchType(request.BatchType);
        var endpointParameter = batchType == "Offboarding" ? "TargetEndpoint" : "SourceEndpoint";
        var lines = new List<string>
        {
            "param(",
            "    [string]$CsvFilePath",
            ")",
            "if (-not [System.IO.File]::Exists($CsvFilePath)) {",
            "    throw \"CSV file not found: $CsvFilePath\"",
            "}",
            "$csvBytes = [System.IO.File]::ReadAllBytes($CsvFilePath)",
            "$params = @{",
            $"    Name = '{EscapePs(request.Name)}'",
            "    CSVData = $csvBytes",
            "}",
            $"$params['{endpointParameter}'] = '{EscapePs(request.EndpointIdentity)}'"
        };

        if (!string.IsNullOrWhiteSpace(request.TargetDeliveryDomain) && batchType != "IMAP")
        {
            lines.Add($"$params['TargetDeliveryDomain'] = '{EscapePs(request.TargetDeliveryDomain)}'");
        }

        var notificationEmails = request.NotificationEmails
            .Where(static value => !string.IsNullOrWhiteSpace(value))
            .Select(static value => value.Trim())
            .ToArray();

        if (notificationEmails.Length > 0)
        {
            lines.Add($"$params['NotificationEmails'] = {ToPsArrayLiteral(notificationEmails)}");
        }

        if (request.AutoStart)
        {
            lines.Add("$params['AutoStart'] = $true");
        }

        if (request.AutoComplete)
        {
            lines.Add("$params['AutoComplete'] = $true");
        }

        lines.Add("New-MigrationBatch @params -ErrorAction Stop | Out-Null");
        lines.Add("Write-Output 'OK'");

        return (string.Join(Environment.NewLine, lines), new Dictionary<string, object>
        {
            ["CsvFilePath"] = request.CsvFilePath
        });
    }

    private static void AddStringParam(List<string> lines, string parameterName, string? value)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            lines.Add($"$params['{parameterName}'] = '{EscapePs(value)}'");
        }
    }

    private static void AddNullableIntParam(List<string> lines, string parameterName, int? value)
    {
        if (value.HasValue)
        {
            lines.Add($"$params['{parameterName}'] = {value.Value}");
        }
    }

    private static bool RequiresCredential(string endpointType)
        => endpointType is "ExchangeRemoteMove" or "ExchangeOutlookAnywhere";

    private static string NormalizeEndpointType(string? value)
        => value?.Trim() switch
        {
            "ExchangeRemoteMove" => "ExchangeRemoteMove",
            "ExchangeOutlookAnywhere" => "ExchangeOutlookAnywhere",
            "IMAP" => "IMAP",
            _ => "ExchangeRemoteMove"
        };

    private static string NormalizeBatchType(string? value)
        => value?.Trim() switch
        {
            "Offboarding" => "Offboarding",
            "IMAP" => "IMAP",
            _ => "Onboarding"
        };

    private static string EscapePs(string? value)
        => (value ?? string.Empty).Replace("'", "''");

    private static string ToPsArrayLiteral(IEnumerable<string> values)
        => "@(" + string.Join(", ", values.Select(static value => $"'{EscapePs(value)}'")) + ")";
}
