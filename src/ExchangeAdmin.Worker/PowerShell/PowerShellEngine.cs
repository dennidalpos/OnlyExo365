using System.Diagnostics;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace ExchangeAdmin.Worker.PowerShell;

/// <summary>
/// Risultato inizializzazione PowerShell.
/// </summary>
public class PowerShellInitResult
{
    /// <summary>True se inizializzazione riuscita.</summary>
    public bool Success { get; init; }

    /// <summary>Versione PowerShell rilevata.</summary>
    public string? PowerShellVersion { get; init; }

    /// <summary>True se ExchangeOnlineManagement module è disponibile.</summary>
    public bool IsModuleAvailable { get; init; }

    /// <summary>Messaggio di errore in caso di fallimento.</summary>
    public string? ErrorMessage { get; init; }
}

/// <summary>
/// Risultato esecuzione comando PowerShell.
/// </summary>
public class PowerShellResult
{
    /// <summary>True se esecuzione completata senza errori.</summary>
    public bool Success { get; init; }

    /// <summary>Output objects dal comando.</summary>
    public List<PSObject> Output { get; init; } = new();

    /// <summary>Errori rilevati durante esecuzione.</summary>
    public List<ErrorRecord> Errors { get; init; } = new();

    /// <summary>Messaggi verbose.</summary>
    public List<string> Verbose { get; init; } = new();

    /// <summary>Messaggi warning.</summary>
    public List<string> Warning { get; init; } = new();

    /// <summary>True se l'operazione è stata cancellata.</summary>
    public bool WasCancelled { get; init; }

    /// <summary>Messaggio di errore principale.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>True se il runspace è in stato corrotto e richiede reset.</summary>
    public bool RunspaceCorrupted { get; init; }
}

/// <summary>
/// Engine PowerShell per l'esecuzione dei comandi Exchange Online.
/// Gestisce runspace lifecycle, pipeline cleanup, e recovery da stati corrotti.
/// </summary>
public sealed class PowerShellEngine : IDisposable
{
    private Runspace? _runspace;
    private bool _isModuleAvailable;
    private string? _powerShellVersion;
    private bool _isInitialized;
    private bool _isConnected;
    private readonly SemaphoreSlim _executionLock = new(1, 1);
    private readonly object _stateLock = new();
    private volatile bool _isDisposing;
    private int _consecutiveFailures;
    private const int MaxConsecutiveFailuresBeforeReset = 3;

    /// <summary>Indica se l'engine è inizializzato.</summary>
    public bool IsInitialized => _isInitialized;

    /// <summary>Indica se è connesso a Exchange Online.</summary>
    public bool IsConnected => _isConnected;

    /// <summary>Indica se il modulo EXO è disponibile.</summary>
    public bool IsModuleAvailable => _isModuleAvailable;

    /// <summary>Versione PowerShell.</summary>
    public string? PowerShellVersion => _powerShellVersion;

    /// <summary>
    /// Inizializza il runspace PowerShell.
    /// </summary>
    public async Task<PowerShellInitResult> InitializeAsync()
    {
        try
        {
            Debug.WriteLine("[PowerShellEngine] Initializing...");

            // Crea runspace con default session state
            var iss = InitialSessionState.CreateDefault();

            // Imposta Execution Policy per permettere script firmati e moduli
            iss.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.RemoteSigned;

            _runspace = RunspaceFactory.CreateRunspace(iss);
            _runspace.Open();

            Debug.WriteLine($"[PowerShellEngine] Runspace opened, state: {_runspace.RunspaceStateInfo.State}");

            // Ottieni versione PowerShell
            using var ps = System.Management.Automation.PowerShell.Create();
            ps.Runspace = _runspace;
            ps.AddScript("$PSVersionTable.PSVersion.ToString()");

            var versionResult = await Task.Run(() => ps.Invoke()).ConfigureAwait(false);
            _powerShellVersion = versionResult.FirstOrDefault()?.ToString() ?? "Unknown";

            Debug.WriteLine($"[PowerShellEngine] PowerShell version: {_powerShellVersion}");

            // Verifica disponibilità modulo ExchangeOnlineManagement
            ps.Commands.Clear();
            ps.AddScript("Get-Module -ListAvailable -Name ExchangeOnlineManagement | Select-Object -First 1");

            var moduleResult = await Task.Run(() => ps.Invoke()).ConfigureAwait(false);
            _isModuleAvailable = moduleResult.Any();

            Debug.WriteLine($"[PowerShellEngine] ExchangeOnlineManagement module available: {_isModuleAvailable}");

            if (_isModuleAvailable)
            {
                // Importa il modulo
                ps.Commands.Clear();
                ps.AddScript("Import-Module ExchangeOnlineManagement -ErrorAction Stop");
                await Task.Run(() => ps.Invoke()).ConfigureAwait(false);

                if (ps.HadErrors)
                {
                    var errors = string.Join("; ", ps.Streams.Error.Select(e => e.ToString()));
                    Debug.WriteLine($"[PowerShellEngine] Warning: Module import had errors: {errors}");
                }
                else
                {
                    Debug.WriteLine("[PowerShellEngine] Module imported successfully");
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

    /// <summary>
    /// Esegue un comando PowerShell con supporto cancellazione e streaming output.
    /// Gestisce cleanup della pipeline e recovery da stati corrotti.
    /// </summary>
    /// <param name="script">Script PowerShell da eseguire.</param>
    /// <param name="parameters">Parametri opzionali.</param>
    /// <param name="onVerbose">Callback per messaggi verbose.</param>
    /// <param name="onWarning">Callback per messaggi warning.</param>
    /// <param name="onError">Callback per errori.</param>
    /// <param name="onOutput">Callback per output objects.</param>
    /// <param name="cancellationToken">Token di cancellazione.</param>
    /// <returns>Risultato dell'esecuzione.</returns>
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

        // Verifica stato runspace prima di acquisire lock
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
#pragma warning disable CA2000 // Dispose handled in finally block below
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

            // Setup stream handlers
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
                    errors.Add(record);
                    onError?.Invoke(record);
                }
            };

            // Registra per cancellazione
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
                // Esegui con output streaming
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

                // Reset failure counter on success
                if (!ps.HadErrors)
                {
                    _consecutiveFailures = 0;
                }

                return new PowerShellResult
                {
                    Success = !ps.HadErrors,
                    Output = output,
                    Errors = errors,
                    Verbose = verbose,
                    Warning = warnings,
                    ErrorMessage = errors.Any() ? string.Join("; ", errors.Select(e => e.ToString())) : null
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
            // Cleanup della PowerShell instance
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

                ps = null; // Indica all'analyzer che è stato gestito
            }

            _executionLock.Release();
        }
    }

    /// <summary>
    /// Verifica se il runspace è in uno stato utilizzabile.
    /// </summary>
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

    /// <summary>
    /// Tenta di recuperare il runspace da uno stato corrotto.
    /// </summary>
    private async Task<bool> TryRecoverRunspaceAsync()
    {
        Debug.WriteLine("[PowerShellEngine] Attempting runspace recovery...");

        try
        {
            // Chiudi vecchio runspace se esiste
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

            // Ricrea runspace
            var iss = InitialSessionState.CreateDefault();
            _runspace = RunspaceFactory.CreateRunspace(iss);
            _runspace.Open();

            // Re-importa modulo se era disponibile
            if (_isModuleAvailable)
            {
                using var ps = System.Management.Automation.PowerShell.Create();
                ps.Runspace = _runspace;
                ps.AddScript("Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue");
                await Task.Run(() => ps.Invoke()).ConfigureAwait(false);
            }

            _consecutiveFailures = 0;
            _isConnected = false; // Reset connection state - richiede riconnessione

            Debug.WriteLine("[PowerShellEngine] Runspace recovery successful");
            return true;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[PowerShellEngine] Runspace recovery failed: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Connetti a Exchange Online (interattivo).
    /// </summary>
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

        var result = await ExecuteAsync(
            "Connect-ExchangeOnline -ShowBanner:$false",
            onVerbose: onVerbose,
            onWarning: (level, msg) => Console.WriteLine($"[PS Warning] {msg}"),
            onError: (err) => Console.WriteLine($"[PS Error] {err}"),
            cancellationToken: cancellationToken).ConfigureAwait(false);

        if (result.Success)
        {
            _isConnected = true;
            Console.WriteLine("[PowerShellEngine] Connected to Exchange Online");
            Debug.WriteLine("[PowerShellEngine] Connected to Exchange Online");
        }
        else
        {
            Console.WriteLine($"[PowerShellEngine] Connection failed: {result.ErrorMessage}");
            Debug.WriteLine($"[PowerShellEngine] Connection failed: {result.ErrorMessage}");

            // Log dettagli degli errori
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

    /// <summary>
    /// Disconnetti da Exchange Online.
    /// </summary>
    public async Task<PowerShellResult> DisconnectExchangeAsync(CancellationToken cancellationToken = default)
    {
        Debug.WriteLine("[PowerShellEngine] Disconnecting from Exchange Online...");

        var result = await ExecuteAsync(
            "Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue",
            cancellationToken: cancellationToken).ConfigureAwait(false);

        _isConnected = false;

        Debug.WriteLine("[PowerShellEngine] Disconnected from Exchange Online");

        return result;
    }

    /// <summary>
    /// Verifica connessione attiva.
    /// </summary>
    public async Task<(bool IsConnected, string? UserPrincipalName, string? Organization)> GetConnectionStatusAsync(
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

                    return (isConnected, upn, org);
                }
            }

            lock (_stateLock)
            {
                _isConnected = false;
            }

            return (false, null, null);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[PowerShellEngine] GetConnectionStatus error: {ex.Message}");

            lock (_stateLock)
            {
                _isConnected = false;
            }

            return (false, null, null);
        }
    }

    /// <summary>
    /// Rilascia le risorse.
    /// </summary>
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
