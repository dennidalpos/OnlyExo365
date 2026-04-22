using System.Diagnostics;
using OnlyExo365.Contracts;

namespace OnlyExo365.Worker.PowerShell;

internal sealed class ExchangeConnectionSession
{
    private readonly PowerShellEngine _engine;
    private readonly PowerShellExecutionPipeline _executionPipeline;

    public ExchangeConnectionSession(PowerShellEngine engine, PowerShellExecutionPipeline executionPipeline)
    {
        _engine = engine;
        _executionPipeline = executionPipeline;
    }

    public async Task<PowerShellResult> ConnectExchangeAsync(
        Action<string, string>? onVerbose = null,
        CancellationToken cancellationToken = default)
    {
        if (!_engine.ModuleAvailable)
        {
            return new PowerShellResult
            {
                Success = false,
                ErrorMessage = "ExchangeOnlineManagement module is not available. Please install it using: Install-Module ExchangeOnlineManagement"
            };
        }

        var validationErrors = _engine.ExchangeConfiguration.Validate();
        if (validationErrors.Count > 0)
        {
            return new PowerShellResult
            {
                Success = false,
                ErrorMessage = string.Join(" ", validationErrors)
            };
        }

        ConsoleLogger.Info("PowerShellEngine", $"Connecting to Exchange Online using {_engine.ExchangeConfiguration.AuthenticationMode}...");
        Debug.WriteLine($"[PowerShellEngine] Connecting to Exchange Online using {_engine.ExchangeConfiguration.AuthenticationMode}...");

        var connectCommand = ExchangeCommandBuilder.BuildConnectExchangeCommand(_engine.ExchangeConfiguration);

        await _executionPipeline.ExecuteAsync(
            "Disconnect-MgGraph -ErrorAction SilentlyContinue",
            cancellationToken: cancellationToken).ConfigureAwait(false);
        _engine.GraphConnected = false;
        _engine.ConnectedGraphScopesValue = Array.Empty<string>();

        var result = await _executionPipeline.ExecuteAsync(
            connectCommand,
            onVerbose: onVerbose,
            onWarning: static (_, msg) => ConsoleLogger.Warning("PowerShellEngine", msg),
            onError: static err => ConsoleLogger.Error("PowerShellEngine", err.Exception?.Message ?? err.ToString()),
            cancellationToken: cancellationToken).ConfigureAwait(false);

        if (result.Success)
        {
            _engine.Connected = true;
            ConsoleLogger.Success("PowerShellEngine", "Connected to Exchange Online");
            Debug.WriteLine("[PowerShellEngine] Connected to Exchange Online");

            await ConnectInitialServiceBundleAsync(onVerbose, cancellationToken).ConfigureAwait(false);
        }
        else
        {
            ConsoleLogger.Error("PowerShellEngine", $"Connection failed: {result.ErrorMessage}");
            Debug.WriteLine($"[PowerShellEngine] Connection failed: {result.ErrorMessage}");

            if (result.Errors.Count > 0)
            {
                foreach (var err in result.Errors)
                {
                    ConsoleLogger.Error("PowerShellEngine", $"Error detail: {err.Exception?.Message ?? err.ToString()}");
                }
            }
        }

        return result;
    }

    public async Task<PowerShellResult> DisconnectExchangeAsync(CancellationToken cancellationToken = default)
    {
        Debug.WriteLine("[PowerShellEngine] Disconnecting from Exchange Online...");

        try
        {
            await _executionPipeline.ExecuteAsync(
                "Disconnect-MgGraph -ErrorAction SilentlyContinue",
                cancellationToken: cancellationToken).ConfigureAwait(false);
            Debug.WriteLine("[PowerShellEngine] Disconnected from Microsoft Graph");
            _engine.GraphConnected = false;
            _engine.ConnectedGraphScopesValue = Array.Empty<string>();
        }
        catch
        {
        }

        var result = await _executionPipeline.ExecuteAsync(
            "Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue",
            cancellationToken: cancellationToken).ConfigureAwait(false);

        _engine.Connected = false;
        _engine.ComplianceConnected = false;

        Debug.WriteLine("[PowerShellEngine] Disconnected from Exchange Online");

        return result;
    }

    public async Task<(bool IsConnected, string? UserPrincipalName, string? Organization, bool IsGraphConnected, bool IsComplianceConnected)> GetConnectionStatusAsync(
        CancellationToken cancellationToken = default)
    {
        try
        {
            var result = await _executionPipeline.ExecuteAsync(
                @"
                try {
                    $conn = Get-ConnectionInformation -ErrorAction Stop | Select-Object -First 1
                    if ($conn) {
                        @{
                            IsConnected = $true
                            UserPrincipalName = $conn.UserPrincipalName
                            Organization = $conn.Organization
                        }
                    } else {
                        @{ IsConnected = $false }
                    }
                } catch {
                    @{ IsConnected = $false }
                }
                ",
                cancellationToken: cancellationToken).ConfigureAwait(false);

            if (result.Success && result.Output.Any())
            {
                var output = result.Output.First();
                var dict = output.BaseObject as System.Collections.Hashtable;

                if (dict != null)
                {
                    var isConnected = dict["IsConnected"] as bool? ?? false;
                    var upn = dict["UserPrincipalName"]?.ToString();
                    var org = dict["Organization"]?.ToString();

                    lock (_engine.StateLock)
                    {
                        _engine.Connected = isConnected;
                    }

                    return (isConnected, upn, org, _engine.GraphConnected, _engine.ComplianceConnected);
                }
            }

            lock (_engine.StateLock)
            {
                _engine.Connected = false;
            }

            return (false, null, null, _engine.GraphConnected, _engine.ComplianceConnected);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[PowerShellEngine] GetConnectionStatus error: {ex.Message}");

            lock (_engine.StateLock)
            {
                _engine.Connected = false;
            }

            return (false, null, null, _engine.GraphConnected, _engine.ComplianceConnected);
        }
    }

    private async Task ConnectInitialServiceBundleAsync(
        Action<string, string>? onVerbose,
        CancellationToken cancellationToken)
    {
        var graphResult = await ConnectMicrosoftGraphAsync(
            ignoreAutoConnectConfiguration: true,
            delegatedScopes: _engine.ExchangeConfiguration.NormalizeGraphScopes(),
            onVerbose: onVerbose,
            cancellationToken: cancellationToken).ConfigureAwait(false);

        if (!graphResult.Success || !_engine.GraphConnected)
        {
            var graphMessage = graphResult.ErrorMessage ?? "Microsoft Graph initial connection unavailable.";
            ConsoleLogger.Warning("PowerShellEngine", graphMessage);
            onVerbose?.Invoke("Warning", $"{graphMessage} The worker will retry explicitly if a Graph-backed operation requires it.");
        }

        var complianceResult = await ConnectComplianceAsync(onVerbose, cancellationToken).ConfigureAwait(false);
        if (!complianceResult.Success || !_engine.ComplianceConnected)
        {
            var complianceMessage = complianceResult.ErrorMessage ?? "Compliance PowerShell initial connection unavailable.";
            ConsoleLogger.Warning("PowerShellEngine", complianceMessage);
            onVerbose?.Invoke("Warning", $"{complianceMessage} The worker will retry explicitly if a Compliance cmdlet is required.");
        }
    }

    public async Task<PowerShellResult> ConnectMicrosoftGraphAsync(
        bool ignoreAutoConnectConfiguration = false,
        IEnumerable<string>? delegatedScopes = null,
        Action<string, string>? onVerbose = null,
        CancellationToken cancellationToken = default)
    {
        var usesDelegatedScopes = UsesDelegatedScopes(_engine.ExchangeConfiguration.AuthenticationMode);
        var requiredDelegatedScopes = usesDelegatedScopes
            ? _engine.ExchangeConfiguration.NormalizeGraphScopes(delegatedScopes)
            : Array.Empty<string>();

        if (_engine.GraphConnected)
        {
            if (!usesDelegatedScopes || HasRequiredScopes(_engine.ConnectedGraphScopesValue, requiredDelegatedScopes))
            {
                onVerbose?.Invoke("Information", "Microsoft Graph already connected");
                return new PowerShellResult
                {
                    Success = true,
                    Output = { System.Management.Automation.PSObject.AsPSObject("Microsoft Graph already connected") }
                };
            }

            onVerbose?.Invoke("Information", "Microsoft Graph connected with insufficient delegated scopes. Reconnecting...");
        }

        ConsoleLogger.Info("PowerShellEngine", $"Connecting to Microsoft Graph using {_engine.ExchangeConfiguration.AuthenticationMode}...");
        Debug.WriteLine($"[PowerShellEngine] Connecting to Microsoft Graph using {_engine.ExchangeConfiguration.AuthenticationMode}...");

        if (!ignoreAutoConnectConfiguration && !_engine.ExchangeConfiguration.EnableGraphAfterExchangeConnect)
        {
            ConsoleLogger.Info("PowerShellEngine", "Microsoft Graph connection disabled by configuration.");
            Debug.WriteLine("[PowerShellEngine] Microsoft Graph connection disabled by configuration.");
            onVerbose?.Invoke("Information", "Microsoft Graph disabled by configuration");
            _engine.GraphConnected = false;
            return new PowerShellResult
            {
                Success = true,
                Warning = { "Microsoft Graph connection disabled by configuration" }
            };
        }

        var graphConnectScript = string.Join(Environment.NewLine, new[]
        {
            "try {",
            "    $graphModule = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Select-Object -First 1",
            "    if (-not $graphModule) {",
            "        Write-Warning 'Microsoft.Graph.Authentication module not installed'",
            "        return",
            "    }",
            string.Empty,
            "    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop",
            "    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null",
            string.Empty,
            $"    {GraphCommandBuilder.BuildConnectGraphCommand(_engine.ExchangeConfiguration, requiredDelegatedScopes)} -ErrorAction Stop",
            string.Empty,
            "    $ctx = Get-MgContext",
            "    if ($ctx) {",
            "        Write-Output \"Connected to Microsoft Graph (TenantId=$($ctx.TenantId), AuthType=$($ctx.AuthType))\"",
            "    } else {",
            "        Write-Output 'Connected to Microsoft Graph'",
            "    }",
            "} catch {",
            "    Write-Warning \"Connect-MgGraph failed: $($_.Exception.Message)\"",
            "}"
        });

        var graphResult = await _executionPipeline.ExecuteAsync(
            graphConnectScript,
            onVerbose: onVerbose,
            onWarning: static (_, msg) => ConsoleLogger.Warning("PowerShellEngine", msg),
            onError: static err => ConsoleLogger.Error("PowerShellEngine", err.Exception?.Message ?? err.ToString()),
            cancellationToken: cancellationToken).ConfigureAwait(false);

        if (graphResult.Success && graphResult.Output.Any())
        {
            ConsoleLogger.Success("PowerShellEngine", "Connected to Microsoft Graph");
            Debug.WriteLine("[PowerShellEngine] Connected to Microsoft Graph");
            onVerbose?.Invoke("Information", "Connected to Microsoft Graph");
            _engine.GraphConnected = true;
            _engine.ConnectedGraphScopesValue = requiredDelegatedScopes.ToList();
            return graphResult;
        }

        ConsoleLogger.Warning("PowerShellEngine", "Microsoft Graph connection failed (license features will be unavailable)");
        Debug.WriteLine("[PowerShellEngine] Microsoft Graph connection failed");
        onVerbose?.Invoke("Warning", "Microsoft Graph connection failed - license features will be unavailable");
        _engine.GraphConnected = false;
        _engine.ConnectedGraphScopesValue = Array.Empty<string>();
        return graphResult;
    }

    public async Task<PowerShellResult> ConnectComplianceAsync(
        Action<string, string>? onVerbose = null,
        CancellationToken cancellationToken = default)
    {
        if (_engine.ComplianceConnected)
        {
            onVerbose?.Invoke("Information", "Compliance PowerShell already connected");
            return new PowerShellResult
            {
                Success = true,
                Output = { System.Management.Automation.PSObject.AsPSObject("Compliance PowerShell already connected") }
            };
        }

        if (!_engine.ModuleAvailable)
        {
            return new PowerShellResult
            {
                Success = false,
                ErrorMessage = "ExchangeOnlineManagement module is not available. Please install it using: Install-Module ExchangeOnlineManagement"
            };
        }

        var connectCommand = ComplianceCommandBuilder.BuildConnectComplianceCommand(_engine.ExchangeConfiguration);
        var result = await _executionPipeline.ExecuteAsync(
            connectCommand,
            onVerbose: onVerbose,
            onWarning: static (_, msg) => ConsoleLogger.Warning("PowerShellEngine", msg),
            onError: static err => ConsoleLogger.Error("PowerShellEngine", err.Exception?.Message ?? err.ToString()),
            cancellationToken: cancellationToken).ConfigureAwait(false);

        if (result.Success)
        {
            _engine.ComplianceConnected = true;
            ConsoleLogger.Success("PowerShellEngine", "Connected to Security & Compliance PowerShell");
        }
        else
        {
            _engine.ComplianceConnected = false;
            ConsoleLogger.Warning("PowerShellEngine", $"Compliance connection failed: {result.ErrorMessage}");
        }

        return result;
    }

    private static bool UsesDelegatedScopes(ExchangeAuthenticationMode authenticationMode)
        => authenticationMode is ExchangeAuthenticationMode.Interactive or ExchangeAuthenticationMode.DeviceCode;

    private static bool HasRequiredScopes(
        IReadOnlyList<string> connectedScopes,
        IReadOnlyList<string> requiredScopes)
    {
        if (requiredScopes.Count == 0)
        {
            return true;
        }

        if (connectedScopes.Count == 0)
        {
            return false;
        }

        var connected = new HashSet<string>(connectedScopes, StringComparer.OrdinalIgnoreCase);
        return requiredScopes.All(connected.Contains);
    }
}

