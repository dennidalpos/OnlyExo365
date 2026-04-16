using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts;
using ExchangeAdmin.Contracts.Diagnostics;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Errors;
using ExchangeAdmin.Domain.Results;

namespace ExchangeAdmin.Presentation.Services;

internal sealed class InteractiveExchangeBootstrapService : IInteractiveExchangeBootstrapService
{
    private static readonly HashSet<string> SupportedExchangeEnvironments = new(StringComparer.OrdinalIgnoreCase)
    {
        "O365Default",
        "O365GermanyCloud",
        "O365USGovGCCHigh",
        "O365USGovDoD",
        "O365China"
    };

    private readonly ExchangeOnlineConfiguration _configuration;

    public InteractiveExchangeBootstrapService(ExchangeOnlineConfiguration configuration)
    {
        _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
    }

    public async Task<Result> EnsureReadyAsync(
        Action<LogLevel, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        if (_configuration.AuthenticationMode != ExchangeAuthenticationMode.Interactive)
        {
            return Result.Success();
        }

        var tempDirectory = Path.Combine(Path.GetTempPath(), "OnlyExo365", "interactive-auth");
        Directory.CreateDirectory(tempDirectory);

        var scriptPath = Path.Combine(tempDirectory, $"bootstrap-{Guid.NewGuid():N}.ps1");
        var statusPath = Path.Combine(tempDirectory, $"bootstrap-{Guid.NewGuid():N}.json");

        try
        {
            await File.WriteAllTextAsync(
                scriptPath,
                BuildBootstrapScript(_configuration),
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                cancellationToken).ConfigureAwait(false);

            onLog?.Invoke(LogLevel.Information, "Opening interactive Exchange sign-in window...");

            using var process = StartBootstrapProcess(scriptPath, statusPath);
            await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);

            var status = await ReadStatusAsync(statusPath, cancellationToken).ConfigureAwait(false);
            if (status == null)
            {
                return Result.Failure(NormalizedError.Create(
                    ErrorCode.AuthenticationFailed,
                    "Interactive sign-in did not produce a result. The authentication host may have been closed before completion."));
            }

            if (!string.IsNullOrWhiteSpace(status.GraphWarning))
            {
                onLog?.Invoke(LogLevel.Warning, status.GraphWarning);
            }

            if (!string.IsNullOrWhiteSpace(status.ComplianceWarning))
            {
                onLog?.Invoke(LogLevel.Warning, status.ComplianceWarning);
            }

            if (!status.Success)
            {
                return Result.Failure(NormalizedError.Create(
                    ErrorCode.AuthenticationFailed,
                    status.Message ?? "Interactive Exchange sign-in failed."));
            }

            if (!string.IsNullOrWhiteSpace(status.UserPrincipalName) ||
                !string.IsNullOrWhiteSpace(status.Organization))
            {
                onLog?.Invoke(
                    LogLevel.Information,
                    $"Interactive sign-in completed for {status.UserPrincipalName ?? "the selected account"}{FormatOrganization(status.Organization)}.");
            }
            else
            {
                onLog?.Invoke(LogLevel.Information, "Interactive sign-in completed.");
            }

            return Result.Success();
        }
        catch (OperationCanceledException)
        {
            return Result.Cancelled();
        }
        catch (Exception ex)
        {
            return Result.Failure(NormalizedError.FromException(ex));
        }
        finally
        {
            TryDeleteFile(scriptPath);
            TryDeleteFile(statusPath);
        }
    }

    internal static string BuildBootstrapScript(ExchangeOnlineConfiguration configuration)
    {
        ArgumentNullException.ThrowIfNull(configuration);

        var script = new StringBuilder();
        script.AppendLine("param([string]$StatusPath)");
        script.AppendLine("$ErrorActionPreference = 'Stop'");
        script.AppendLine("$ProgressPreference = 'SilentlyContinue'");
        script.AppendLine("$status = [ordered]@{");
        script.AppendLine("    Success = $false");
        script.AppendLine("    Message = $null");
        script.AppendLine("    UserPrincipalName = $null");
        script.AppendLine("    Organization = $null");
        script.AppendLine("    GraphWarning = $null");
        script.AppendLine("    ComplianceWarning = $null");
        script.AppendLine("}");
        script.AppendLine("function Save-Status {");
        script.AppendLine("    param([hashtable]$Payload)");
        script.AppendLine("    $directory = Split-Path -Parent $StatusPath");
        script.AppendLine("    if (-not [string]::IsNullOrWhiteSpace($directory)) {");
        script.AppendLine("        New-Item -ItemType Directory -Path $directory -Force | Out-Null");
        script.AppendLine("    }");
        script.AppendLine("    $Payload | ConvertTo-Json -Depth 6 | Set-Content -Path $StatusPath -Encoding UTF8");
        script.AppendLine("}");
        script.AppendLine("try {");
        script.AppendLine("    Import-Module ExchangeOnlineManagement -ErrorAction Stop");
        script.AppendLine($"    {BuildConnectExchangeCommand(configuration)} -ErrorAction Stop | Out-Null");
        script.AppendLine("    try {");
        script.AppendLine("        $connection = Get-ConnectionInformation -ErrorAction Stop | Select-Object -First 1");
        script.AppendLine("        if ($connection) {");
        script.AppendLine("            $status.UserPrincipalName = $connection.UserPrincipalName");
        script.AppendLine("            $status.Organization = $connection.Organization");
        script.AppendLine("        }");
        script.AppendLine("    } catch {");
        script.AppendLine("    }");
        script.AppendLine();
        script.AppendLine("    try {");
        script.AppendLine("        $graphModule = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Select-Object -First 1");
        script.AppendLine("        if ($graphModule) {");
        script.AppendLine("            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop");
        script.AppendLine($"            {BuildConnectGraphCommand(configuration)} -ErrorAction Stop | Out-Null");
        script.AppendLine("        }");
        script.AppendLine("        else {");
        script.AppendLine("            $status.GraphWarning = 'Microsoft Graph bootstrap skipped because Microsoft.Graph.Authentication is not installed.'");
        script.AppendLine("        }");
        script.AppendLine("    } catch {");
        script.AppendLine("        $status.GraphWarning = \"Interactive Microsoft Graph bootstrap failed: $($_.Exception.Message)\"");
        script.AppendLine("    }");
        script.AppendLine();
        script.AppendLine("    try {");
        script.AppendLine($"        {BuildConnectComplianceCommand(configuration)} -ErrorAction Stop | Out-Null");
        script.AppendLine("    } catch {");
        script.AppendLine("        $status.ComplianceWarning = \"Interactive Compliance bootstrap failed: $($_.Exception.Message)\"");
        script.AppendLine("    }");
        script.AppendLine();
        script.AppendLine("    $status.Success = $true");
        script.AppendLine("} catch {");
        script.AppendLine("    $status.Message = $_.Exception.Message");
        script.AppendLine("} finally {");
        script.AppendLine("    Save-Status -Payload $status");
        script.AppendLine("}");

        return script.ToString();
    }

    private static Process StartBootstrapProcess(string scriptPath, string statusPath)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "pwsh",
            Arguments = $"-NoLogo -ExecutionPolicy Bypass -File \"{scriptPath}\" -StatusPath \"{statusPath}\"",
            UseShellExecute = false,
            CreateNoWindow = false,
            WindowStyle = ProcessWindowStyle.Normal
        };

        return Process.Start(startInfo)
            ?? throw new InvalidOperationException("Unable to start the interactive authentication host.");
    }

    private static async Task<BootstrapStatus?> ReadStatusAsync(string statusPath, CancellationToken cancellationToken)
    {
        if (!File.Exists(statusPath))
        {
            return null;
        }

        await using var stream = File.OpenRead(statusPath);
        return await JsonSerializer.DeserializeAsync<BootstrapStatus>(stream, cancellationToken: cancellationToken)
            .ConfigureAwait(false);
    }

    private static void TryDeleteFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
        }
    }

    private static string FormatOrganization(string? organization)
        => string.IsNullOrWhiteSpace(organization) ? string.Empty : $" to {organization}";

    private static string BuildConnectExchangeCommand(ExchangeOnlineConfiguration configuration)
    {
        var commandParts = BuildExchangeBaseCommand(configuration);

        if (!string.IsNullOrWhiteSpace(configuration.UserPrincipalNameHint))
        {
            commandParts.Add($"-UserPrincipalName '{EscapePs(configuration.UserPrincipalNameHint)}'");
        }

        return string.Join(" ", commandParts);
    }

    private static List<string> BuildExchangeBaseCommand(ExchangeOnlineConfiguration configuration)
    {
        var commandParts = new List<string> { "Connect-ExchangeOnline", "-ShowBanner:$false" };

        if (!string.IsNullOrWhiteSpace(configuration.ExchangeEnvironmentName) &&
            SupportedExchangeEnvironments.Contains(configuration.ExchangeEnvironmentName))
        {
            commandParts.Add($"-ExchangeEnvironmentName '{EscapePs(configuration.ExchangeEnvironmentName)}'");
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

    private static string BuildConnectGraphCommand(ExchangeOnlineConfiguration configuration)
    {
        var commandParts = new List<string>
        {
            "Connect-MgGraph",
            $"-Scopes @({string.Join(", ", configuration.NormalizeGraphScopes().Select(scope => $"'{EscapePs(scope)}'"))})",
            "-ContextScope Process",
            "-NoWelcome"
        };

        if (!string.IsNullOrWhiteSpace(configuration.GraphTenantId))
        {
            commandParts.Add($"-TenantId '{EscapePs(configuration.GraphTenantId)}'");
        }

        return string.Join(" ", commandParts);
    }

    private static string BuildConnectComplianceCommand(ExchangeOnlineConfiguration configuration)
    {
        var commandParts = new List<string> { "Connect-IPPSSession", "-WarningAction", "SilentlyContinue" };

        if (!string.IsNullOrWhiteSpace(configuration.ExchangeOrganization))
        {
            commandParts.Add($"-Organization '{EscapePs(configuration.ExchangeOrganization)}'");
        }

        if (!string.IsNullOrWhiteSpace(configuration.DelegatedOrganization))
        {
            commandParts.Add($"-DelegatedOrganization '{EscapePs(configuration.DelegatedOrganization)}'");
        }

        if (!string.IsNullOrWhiteSpace(configuration.UserPrincipalNameHint))
        {
            commandParts.Add($"-UserPrincipalName '{EscapePs(configuration.UserPrincipalNameHint)}'");
        }

        return string.Join(" ", commandParts);
    }

    private static string EscapePs(string value)
        => value.Replace("'", "''", StringComparison.Ordinal);

    private sealed class BootstrapStatus
    {
        public bool Success { get; set; }
        public string? Message { get; set; }
        public string? UserPrincipalName { get; set; }
        public string? Organization { get; set; }
        public string? GraphWarning { get; set; }
        public string? ComplianceWarning { get; set; }
    }
}
