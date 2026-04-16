using ExchangeAdmin.Contracts;

namespace ExchangeAdmin.Worker.PowerShell;

internal sealed class ComplianceSearchOnlyRunner
{
    private readonly ExchangeOnlineConfiguration _configuration;

    public ComplianceSearchOnlyRunner(ExchangeOnlineConfiguration configuration)
    {
        _configuration = configuration?.Clone() ?? throw new ArgumentNullException(nameof(configuration));
    }

    public async Task<PowerShellResult> ExecuteAsync(
        string script,
        string requiredCmdletName,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(script);
        ArgumentException.ThrowIfNullOrWhiteSpace(requiredCmdletName);

        using var engine = new PowerShellEngine(_configuration.Clone());

        onLog?.Invoke("Information", "Initializing dedicated Compliance search-only runner...");
        var initResult = await engine.InitializeAsync().ConfigureAwait(false);
        if (!initResult.Success)
        {
            return new PowerShellResult
            {
                Success = false,
                ErrorMessage = initResult.ErrorMessage ?? "Unable to initialize dedicated Compliance search-only runner."
            };
        }

        if (!initResult.IsModuleAvailable)
        {
            return new PowerShellResult
            {
                Success = false,
                ErrorMessage = "ExchangeOnlineManagement module is not available. Please install it using: Install-Module ExchangeOnlineManagement"
            };
        }

        onLog?.Invoke("Information", "Connecting dedicated Compliance search-only session...");
        var connectResult = await engine.ExecuteAsync(
            ComplianceCommandBuilder.BuildConnectComplianceSearchOnlyCommand(_configuration),
            onVerbose: onLog,
            onWarning: onLog,
            onError: static err => ConsoleLogger.Error("PowerShellEngine", err.Exception?.Message ?? err.ToString()),
            cancellationToken: cancellationToken).ConfigureAwait(false);

        if (!connectResult.Success || connectResult.WasCancelled)
        {
            return connectResult;
        }

        var escapedCmdletName = EscapePs(requiredCmdletName);
        var validationScript = $@"
if (-not (Get-Command -Name '{escapedCmdletName}' -ErrorAction SilentlyContinue)) {{
    throw 'Cmdlet {escapedCmdletName} is not available in the current Security & Compliance PowerShell session.'
}}";

        var validationResult = await engine.ExecuteAsync(
            validationScript,
            onVerbose: onLog,
            onWarning: onLog,
            onError: static err => ConsoleLogger.Error("PowerShellEngine", err.Exception?.Message ?? err.ToString()),
            cancellationToken: cancellationToken).ConfigureAwait(false);

        if (!validationResult.Success || validationResult.WasCancelled)
        {
            return validationResult;
        }

        return await engine.ExecuteAsync(
            script,
            onVerbose: onLog,
            onWarning: onLog,
            onError: static err => ConsoleLogger.Error("PowerShellEngine", err.Exception?.Message ?? err.ToString()),
            cancellationToken: cancellationToken).ConfigureAwait(false);
    }

    private static string EscapePs(string value)
        => value.Replace("'", "''", StringComparison.Ordinal);
}
