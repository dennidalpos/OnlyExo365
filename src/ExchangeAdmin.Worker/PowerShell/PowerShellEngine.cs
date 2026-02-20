using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace ExchangeAdmin.Worker.PowerShell;

             
                                          
              
public class PowerShellInitResult
{
                                                             
    public bool Success { get; init; }

                                                        
    public string? PowerShellVersion { get; init; }

                                                                                  
    public bool IsModuleAvailable { get; init; }

                                                                     
    public string? ErrorMessage { get; init; }
}

             
                                            
              
public class PowerShellResult
{
                                                                      
    public bool Success { get; init; }

                                                      
    public List<PSObject> Output { get; init; } = new();

                                                              
    public List<ErrorRecord> Errors { get; init; } = new();

                                            
    public List<string> Verbose { get; init; } = new();

                                            
    public List<string> Warning { get; init; } = new();

                                                                    
    public bool WasCancelled { get; init; }

                                                          
    public string? ErrorMessage { get; init; }

                                                                                     
    public bool RunspaceCorrupted { get; init; }
}

             
                                                                   
                                                                                
              
public sealed class PowerShellEngine : IDisposable
{
    private const string ExchangeEnvironmentVariable = "EXCHANGEADMIN_EXO_ENV";
    private const string ExchangeOnlineModuleName = "ExchangeOnlineManagement";
    private static readonly HashSet<string> SupportedExchangeEnvironments = new(StringComparer.OrdinalIgnoreCase)
    {
        "O365Default",
        "O365GermanyCloud",
        "O365USGovGCCHigh",
        "O365USGovDoD",
        "O365China"
    };
    private Runspace? _runspace;
    private bool _isModuleAvailable;
    private string? _powerShellVersion;
    private bool _isInitialized;
    private bool _isConnected;
    private bool _isGraphConnected;
    private readonly SemaphoreSlim _executionLock = new(1, 1);
    private readonly object _stateLock = new();
    private volatile bool _isDisposing;
    private int _consecutiveFailures;
    private const int MaxConsecutiveFailuresBeforeReset = 3;

                                                               
    public bool IsInitialized => _isInitialized;

                                                                   
    public bool IsConnected => _isConnected;

    public bool IsGraphConnected => _isGraphConnected;

                                                                  
    public bool IsModuleAvailable => _isModuleAvailable;

                                               
    public string? PowerShellVersion => _powerShellVersion;

                 
                                           
                  
    public async Task<PowerShellInitResult> InitializeAsync()
    {
        try
        {
            Debug.WriteLine("[PowerShellEngine] Initializing...");

                                                      
            var iss = InitialSessionState.CreateDefault();

                                                                              
            iss.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.RemoteSigned;

            _runspace = RunspaceFactory.CreateRunspace(iss);
            _runspace.Open();

            Debug.WriteLine($"[PowerShellEngine] Runspace opened, state: {_runspace.RunspaceStateInfo.State}");

                                          
            using var ps = System.Management.Automation.PowerShell.Create();
            ps.Runspace = _runspace;
            ps.AddScript("$PSVersionTable.PSVersion.ToString()");

            var versionResult = await Task.Run(() => ps.Invoke()).ConfigureAwait(false);
            _powerShellVersion = versionResult.FirstOrDefault()?.ToString() ?? "Unknown";

            Debug.WriteLine($"[PowerShellEngine] PowerShell version: {_powerShellVersion}");

                                                                      
            ps.Commands.Clear();
            ps.AddScript($"Get-Module -ListAvailable -Name {ExchangeOnlineModuleName} | Select-Object -First 1");

            var moduleResult = await Task.Run(() => ps.Invoke()).ConfigureAwait(false);
            _isModuleAvailable = moduleResult.Any();

            Debug.WriteLine($"[PowerShellEngine] ExchangeOnlineManagement module available: {_isModuleAvailable}");

            if (_isModuleAvailable)
            {
                var packageManagementReady = await EnsurePackageManagementAvailableAsync().ConfigureAwait(false);
                if (!packageManagementReady)
                {
                    Debug.WriteLine("[PowerShellEngine] PackageManagement module was not explicitly loaded. Continuing with best-effort import.");
                }

                var imported = await ImportModuleAsync(ExchangeOnlineModuleName, stopOnError: true).ConfigureAwait(false);
                if (!imported)
                {
                    Debug.WriteLine($"[PowerShellEngine] Warning: {ExchangeOnlineModuleName} import reported warnings/errors");
                }
            }

            _isInitialized = true;
            _consecutiveFailures = 0;

            return new PowerShellInitResult
            {
                Success = true,
                PowerShellVersion = _powerShellVersion,
                IsModuleAvailable = _isModuleAvailable
            };
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[PowerShellEngine] Initialization failed: {ex.Message}");

            return new PowerShellInitResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                PowerShellVersion = _powerShellVersion,
                IsModuleAvailable = false
            };
        }
    }

                 
                                                                                   
                                                                     
                  
                                                                   
                                                             
                                                                      
                                                                      
                                                          
                                                                   
                                                                       
                                                     
    public async Task<PowerShellResult> ExecuteAsync(
        string script,
        Dictionary<string, object>? parameters = null,
        Action<string, string>? onVerbose = null,
        Action<string, string>? onWarning = null,
        Action<ErrorRecord>? onError = null,
        Action<PSObject>? onOutput = null,
        CancellationToken cancellationToken = default)
    {
        if (_isDisposing)
        {
            return new PowerShellResult
            {
                Success = false,
                ErrorMessage = "PowerShell engine is disposing"
            };
        }

        if (!_isInitialized || _runspace == null)
        {
            return new PowerShellResult
            {
                Success = false,
                ErrorMessage = "PowerShell engine not initialized"
            };
        }

                                                          
        if (!IsRunspaceUsable())
        {
            Debug.WriteLine("[PowerShellEngine] Runspace not usable, attempting recovery...");

            var recovered = await TryRecoverRunspaceAsync().ConfigureAwait(false);
            if (!recovered)
            {
                return new PowerShellResult
                {
                    Success = false,
                    ErrorMessage = "Runspace is in invalid state and recovery failed",
                    RunspaceCorrupted = true
                };
            }
        }

        await _executionLock.WaitAsync(cancellationToken).ConfigureAwait(false);

        System.Management.Automation.PowerShell? ps = null;

        try
        {
#pragma warning disable CA2000                                          
            ps = System.Management.Automation.PowerShell.Create();
#pragma warning restore CA2000
            ps.Runspace = _runspace;

            ps.AddScript(script);

            if (parameters != null)
            {
                foreach (var param in parameters)
                {
                    ps.AddParameter(param.Key, param.Value);
                }
            }

                                    
            var verbose = new List<string>();
            var warnings = new List<string>();
            var errors = new List<ErrorRecord>();
            var output = new List<PSObject>();

            ps.Streams.Verbose.DataAdded += (s, e) =>
            {
                if (e.Index >= 0 && e.Index < ps.Streams.Verbose.Count)
                {
                    var record = ps.Streams.Verbose[e.Index];
                    var message = record.Message;
                    verbose.Add(message);
                    onVerbose?.Invoke("Verbose", message);
                }
            };

            ps.Streams.Warning.DataAdded += (s, e) =>
            {
                if (e.Index >= 0 && e.Index < ps.Streams.Warning.Count)
                {
                    var record = ps.Streams.Warning[e.Index];
                    var message = record.Message;
                    warnings.Add(message);
                    onWarning?.Invoke("Warning", message);
                }
            };

            ps.Streams.Error.DataAdded += (s, e) =>
            {
                if (e.Index >= 0 && e.Index < ps.Streams.Error.Count)
                {
                    var record = ps.Streams.Error[e.Index];

                    // Check if this is a deprecation warning, treat it as a warning instead of error
                    var errorMessage = record.Exception?.Message ?? record.ToString();
                    if (errorMessage.Contains("deprecat", StringComparison.OrdinalIgnoreCase) ||
                        errorMessage.Contains("will start deprecating", StringComparison.OrdinalIgnoreCase))
                    {
                        // Treat deprecation notices as warnings, not errors
                        warnings.Add(errorMessage);
                        onWarning?.Invoke("Warning", errorMessage);
                    }
                    else
                    {
                        errors.Add(record);
                        onError?.Invoke(record);
                    }
                }
            };

                                         
            var registration = cancellationToken.Register(() =>
            {
                try
                {
                    Debug.WriteLine("[PowerShellEngine] Cancellation requested, stopping pipeline...");
                    ps.Stop();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[PowerShellEngine] Error stopping pipeline: {ex.Message}");
                }
            });

            try
            {
                                              
                Console.WriteLine($"[PowerShellEngine] Starting script execution (length: {script.Length} chars)");
                var scriptPreview = script.Length > 100 ? script.Substring(0, 100) + "..." : script;
                Console.WriteLine($"[PowerShellEngine] Script preview: {scriptPreview.Replace("\n", " ").Replace("\r", "")}");

                await Task.Run(() =>
                {
                    var collection = new PSDataCollection<PSObject>();
                    collection.DataAdded += (s, e) =>
                    {
                        if (e.Index >= 0 && e.Index < collection.Count)
                        {
                            var item = collection[e.Index];
                            output.Add(item);
                            onOutput?.Invoke(item);
                        }
                    };

                    Console.WriteLine($"[PowerShellEngine] Invoking PowerShell...");
                    ps.Invoke(null, collection);
                    Console.WriteLine($"[PowerShellEngine] PowerShell invocation completed. HadErrors: {ps.HadErrors}, Output count: {output.Count}");
                }, cancellationToken).ConfigureAwait(false);

                if (cancellationToken.IsCancellationRequested)
                {
                    Debug.WriteLine("[PowerShellEngine] Execution cancelled");

                    return new PowerShellResult
                    {
                        Success = false,
                        WasCancelled = true,
                        Output = output,
                        Errors = errors,
                        Verbose = verbose,
                        Warning = warnings
                    };
                }

                                                   
                var hasErrors = errors.Count > 0;

                if (!hasErrors)
                {
                    _consecutiveFailures = 0;
                }

                return new PowerShellResult
                {
                    Success = !hasErrors,
                    Output = output,
                    Errors = errors,
                    Verbose = verbose,
                    Warning = warnings,
                    ErrorMessage = hasErrors ? string.Join("; ", errors.Select(e => e.ToString())) : null
                };
            }
            finally
            {
                await registration.DisposeAsync().ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException)
        {
            Debug.WriteLine("[PowerShellEngine] Operation cancelled");

            return new PowerShellResult
            {
                Success = false,
                WasCancelled = true
            };
        }
        catch (PSInvalidOperationException ex)
        {
            Debug.WriteLine($"[PowerShellEngine] PSInvalidOperationException: {ex.Message}");
            _consecutiveFailures++;

            return new PowerShellResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                RunspaceCorrupted = _consecutiveFailures >= MaxConsecutiveFailuresBeforeReset
            };
        }
        catch (InvalidRunspaceStateException ex)
        {
            Debug.WriteLine($"[PowerShellEngine] InvalidRunspaceStateException: {ex.Message}");
            _consecutiveFailures++;

            return new PowerShellResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                RunspaceCorrupted = true
            };
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[PowerShellEngine] Execution error: {ex.GetType().Name} - {ex.Message}");
            _consecutiveFailures++;

            return new PowerShellResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                RunspaceCorrupted = _consecutiveFailures >= MaxConsecutiveFailuresBeforeReset
            };
        }
        finally
        {
                                                
            if (ps != null)
            {
                try
                {
                    ps.Dispose();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[PowerShellEngine] Error disposing PowerShell: {ex.Message}");
                }

                ps = null;                                            
            }

            _executionLock.Release();
        }
    }

                 
                                                             
                  
    private bool IsRunspaceUsable()
    {
        if (_runspace == null)
        {
            return false;
        }

        try
        {
            var state = _runspace.RunspaceStateInfo.State;
            return state == RunspaceState.Opened;
        }
        catch
        {
            return false;
        }
    }

                 
                                                              
                  

    private async Task<bool> ImportModuleAsync(string moduleName, bool stopOnError)
    {
        if (_runspace == null || string.IsNullOrWhiteSpace(moduleName))
        {
            return false;
        }

        using var ps = System.Management.Automation.PowerShell.Create();
        ps.Runspace = _runspace;

        var safeModuleName = moduleName.Replace("'", "''", StringComparison.Ordinal);
        var errorAction = stopOnError ? "Stop" : "SilentlyContinue";
        ps.AddScript($@"
$moduleName = '{safeModuleName}'
$available = Get-Module -ListAvailable -Name $moduleName | Select-Object -First 1
if (-not $available) {{
    Write-Error ""Module '$moduleName' is not installed or not available in PSModulePath.""
}}
else {{
    Import-Module -Name $moduleName -Global -ErrorAction {errorAction} | Out-Null
}}
");

        try
        {
            await Task.Run(() => ps.Invoke()).ConfigureAwait(false);
        }
        catch (RuntimeException ex)
        {
            Debug.WriteLine($"[PowerShellEngine] Module import exception for '{moduleName}': {ex.Message}");
            return false;
        }

        if (ps.HadErrors)
        {
            var errors = string.Join("; ", ps.Streams.Error.Select(e => e.ToString()));
            Debug.WriteLine($"[PowerShellEngine] Module import warning for '{moduleName}': {errors}");
            ps.Streams.Error.Clear();
            return false;
        }

        Debug.WriteLine($"[PowerShellEngine] Module imported: {moduleName}");
        return true;
    }

    private async Task<bool> EnsurePackageManagementAvailableAsync()
    {
        if (_runspace == null)
        {
            return false;
        }

        using var ps = System.Management.Automation.PowerShell.Create();
        ps.Runspace = _runspace;

        ps.AddScript(@"
$requiredVersion = [Version]'1.4.4'
$module = Get-Module -ListAvailable -Name PackageManagement |
    Sort-Object -Property Version -Descending |
    Select-Object -First 1

if (-not $module) {
    Write-Verbose 'PackageManagement module was not found.'
    return $false
}

if ($module.Version -lt $requiredVersion) {
    Write-Verbose ""PackageManagement version $($module.Version) is lower than required $requiredVersion.""
    return $false
}

Import-Module -Name $module.Path -Global -ErrorAction Stop | Out-Null
return $true
");

        try
        {
            var result = await Task.Run(() => ps.Invoke()).ConfigureAwait(false);
            return result.FirstOrDefault()?.BaseObject as bool? ?? false;
        }
        catch (RuntimeException ex)
        {
            Debug.WriteLine($"[PowerShellEngine] Failed to prepare PackageManagement: {ex.Message}");
            return false;
        }
    }

    private async Task<bool> TryRecoverRunspaceAsync()
    {
        Debug.WriteLine("[PowerShellEngine] Attempting runspace recovery...");

        try
        {
                                                
            if (_runspace != null)
            {
                try
                {
                    _runspace.Close();
                    _runspace.Dispose();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[PowerShellEngine] Error closing old runspace: {ex.Message}");
                }
                _runspace = null;
            }

                              
            var iss = InitialSessionState.CreateDefault();
            _runspace = RunspaceFactory.CreateRunspace(iss);
            _runspace.Open();

                                                   
            if (_isModuleAvailable)
            {
                _ = await ImportModuleAsync(ExchangeOnlineModuleName, stopOnError: false).ConfigureAwait(false);
            }

            _consecutiveFailures = 0;
            _isConnected = false;                                                   

            Debug.WriteLine("[PowerShellEngine] Runspace recovery successful");
            return true;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[PowerShellEngine] Runspace recovery failed: {ex.Message}");
            return false;
        }
    }

                 
                                                 
                  
    public async Task<PowerShellResult> ConnectExchangeInteractiveAsync(
        Action<string, string>? onVerbose = null,
        CancellationToken cancellationToken = default)
    {
        if (!_isModuleAvailable)
        {
            return new PowerShellResult
            {
                Success = false,
                ErrorMessage = "ExchangeOnlineManagement module is not available. Please install it using: Install-Module ExchangeOnlineManagement"
            };
        }

        Console.WriteLine("[PowerShellEngine] Connecting to Exchange Online...");
        Debug.WriteLine("[PowerShellEngine] Connecting to Exchange Online...");

        var exchangeEnvironment = Environment.GetEnvironmentVariable(ExchangeEnvironmentVariable);
        var connectCommand = BuildConnectExchangeCommand(exchangeEnvironment);

        var result = await ExecuteAsync(
            connectCommand,
            onVerbose: onVerbose,
            onWarning: (level, msg) => Console.WriteLine($"[PS Warning] {msg}"),
            onError: (err) => Console.WriteLine($"[PS Error] {err}"),
            cancellationToken: cancellationToken).ConfigureAwait(false);

        if (result.Success)
        {
            _isConnected = true;
            Console.WriteLine("[PowerShellEngine] Connected to Exchange Online");
            Debug.WriteLine("[PowerShellEngine] Connected to Exchange Online");

            // Attempt to connect to Microsoft Graph for license and admin operations
            await ConnectMicrosoftGraphAsync(onVerbose, cancellationToken).ConfigureAwait(false);
        }
        else
        {
            Console.WriteLine($"[PowerShellEngine] Connection failed: {result.ErrorMessage}");
            Debug.WriteLine($"[PowerShellEngine] Connection failed: {result.ErrorMessage}");

            if (result.Errors != null && result.Errors.Any())
            {
                foreach (var err in result.Errors)
                {
                    Console.WriteLine($"[PowerShellEngine] Error detail: {err.Exception?.Message ?? err.ToString()}");
                }
            }
        }

        return result;
    }

    private async Task ConnectMicrosoftGraphAsync(
        Action<string, string>? onVerbose = null,
        CancellationToken cancellationToken = default)
    {
        Console.WriteLine("[PowerShellEngine] Connecting to Microsoft Graph...");
        Debug.WriteLine("[PowerShellEngine] Connecting to Microsoft Graph...");

        var graphConnectScript = @"
try {
    $graphModule = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Select-Object -First 1
    if (-not $graphModule) {
        Write-Warning 'Microsoft.Graph.Authentication module not installed'
        return
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    # Use interactive authentication (same as Exchange) with required scopes
    # RoleManagement.Read.All is required for directory role membership queries
    # User.ReadWrite.All is required for license assignment/removal
    $scopes = @('Organization.Read.All', 'Directory.Read.All', 'RoleManagement.Read.All', 'User.ReadWrite.All')
    Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
    Write-Output 'Connected to Microsoft Graph'
} catch {
    Write-Warning ""Connect-MgGraph failed: $($_.Exception.Message)""
}";

        var graphResult = await ExecuteAsync(
            graphConnectScript,
            onVerbose: onVerbose,
            onWarning: (level, msg) => Console.WriteLine($"[PS Warning] {msg}"),
            onError: (err) => Console.WriteLine($"[PS Error] {err}"),
            cancellationToken: cancellationToken).ConfigureAwait(false);

        if (graphResult.Success && graphResult.Output.Any())
        {
            Console.WriteLine("[PowerShellEngine] Connected to Microsoft Graph");
            Debug.WriteLine("[PowerShellEngine] Connected to Microsoft Graph");
            onVerbose?.Invoke("Information", "Connected to Microsoft Graph");
            _isGraphConnected = true;
        }
        else
        {
            Console.WriteLine("[PowerShellEngine] Microsoft Graph connection failed (license features will be unavailable)");
            Debug.WriteLine("[PowerShellEngine] Microsoft Graph connection failed");
            onVerbose?.Invoke("Warning", "Microsoft Graph connection failed - license features will be unavailable");
            _isGraphConnected = false;
        }
    }

    private static string BuildConnectExchangeCommand(string? exchangeEnvironment)
    {
        const string baseCommand = "Connect-ExchangeOnline -ShowBanner:$false";
        if (string.IsNullOrWhiteSpace(exchangeEnvironment))
        {
            return baseCommand;
        }

        if (!SupportedExchangeEnvironments.Contains(exchangeEnvironment))
        {
            Console.WriteLine($"[PowerShellEngine] Unsupported Exchange environment '{exchangeEnvironment}', falling back to default.");
            return baseCommand;
        }

        var sanitized = exchangeEnvironment.Replace("'", "''", StringComparison.Ordinal);
        return $"{baseCommand} -ExchangeEnvironmentName '{sanitized}'";
    }

                 
                                       
                  
    public async Task<PowerShellResult> DisconnectExchangeAsync(CancellationToken cancellationToken = default)
    {
        Debug.WriteLine("[PowerShellEngine] Disconnecting from Exchange Online...");

        // Disconnect from Microsoft Graph first
        try
        {
            await ExecuteAsync(
                "Disconnect-MgGraph -ErrorAction SilentlyContinue",
                cancellationToken: cancellationToken).ConfigureAwait(false);
            Debug.WriteLine("[PowerShellEngine] Disconnected from Microsoft Graph");
            _isGraphConnected = false;
        }
        catch
        {
            // Ignore Graph disconnect errors
        }

        var result = await ExecuteAsync(
            "Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue",
            cancellationToken: cancellationToken).ConfigureAwait(false);

        _isConnected = false;

        Debug.WriteLine("[PowerShellEngine] Disconnected from Exchange Online");

        return result;
    }

                 
                                    
                  
    public async Task<(bool IsConnected, string? UserPrincipalName, string? Organization, bool IsGraphConnected)> GetConnectionStatusAsync(
        CancellationToken cancellationToken = default)
    {
        try
        {
            var result = await ExecuteAsync(
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

                    lock (_stateLock)
                    {
                        _isConnected = isConnected;
                    }

                    return (isConnected, upn, org, _isGraphConnected);
                }
            }

            lock (_stateLock)
            {
                _isConnected = false;
            }

            return (false, null, null, _isGraphConnected);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[PowerShellEngine] GetConnectionStatus error: {ex.Message}");

            lock (_stateLock)
            {
                _isConnected = false;
            }

            return (false, null, null, _isGraphConnected);
        }
    }

                 
                            
                  
    public void Dispose()
    {
        if (_isDisposing)
        {
            return;
        }

        _isDisposing = true;

        Debug.WriteLine("[PowerShellEngine] Disposing...");

        try
        {
            _executionLock.Dispose();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[PowerShellEngine] Error disposing lock: {ex.Message}");
        }

        try
        {
            _runspace?.Close();
            _runspace?.Dispose();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[PowerShellEngine] Error disposing runspace: {ex.Message}");
        }

        Debug.WriteLine("[PowerShellEngine] Disposed");
    }
}
