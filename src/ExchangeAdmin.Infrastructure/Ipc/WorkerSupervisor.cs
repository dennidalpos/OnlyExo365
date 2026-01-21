using System.Diagnostics;
using ExchangeAdmin.Contracts;
using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Infrastructure.Ipc;

/// <summary>
/// Configurazione del supervisor per il worker PowerShell.
/// </summary>
public class WorkerSupervisorOptions
{
    /// <summary>
    /// Path all'eseguibile del worker. Può essere relativo o assoluto.
    /// </summary>
    public string WorkerPath { get; set; } = "ExchangeAdmin.Worker.exe";

    /// <summary>
    /// Numero massimo di restart automatici prima di arrendersi.
    /// </summary>
    public int MaxRestartAttempts { get; set; } = 3;

    /// <summary>
    /// Intervallo minimo tra restart (ms). Evita restart loop.
    /// </summary>
    public int RestartCooldownMs { get; set; } = 2000;

    /// <summary>
    /// Timeout attesa avvio worker e handshake (ms).
    /// </summary>
    public int StartupTimeoutMs { get; set; } = 30000;

    /// <summary>
    /// Intervallo tra heartbeat ping (ms).
    /// </summary>
    public int HeartbeatIntervalMs { get; set; } = IpcConstants.HeartbeatIntervalMs;

    /// <summary>
    /// Timeout per singolo heartbeat - dopo questo tempo senza pong, incrementa missed count (ms).
    /// </summary>
    public int HeartbeatTimeoutMs { get; set; } = IpcConstants.HeartbeatTimeoutMs;

    /// <summary>
    /// Grace period dopo heartbeat timeout prima di dichiarare worker morto (ms).
    /// Permette retry aggiuntivi.
    /// </summary>
    public int HeartbeatGracePeriodMs { get; set; } = IpcConstants.HeartbeatGracePeriodMs;

    /// <summary>
    /// Numero di heartbeat consecutivi mancati prima di kill.
    /// </summary>
    public int HeartbeatMissedThreshold { get; set; } = IpcConstants.HeartbeatMissedThreshold;

    /// <summary>
    /// Ambiente Exchange Online per l'autenticazione (es. O365Default).
    /// Se null/empty usa il valore di default del modulo.
    /// </summary>
    public string? ExchangeEnvironmentName { get; set; }
}

/// <summary>
/// Informazioni sullo stato corrente del worker.
/// </summary>
public class WorkerStatus
{
    /// <summary>Stato connessione corrente.</summary>
    public WorkerConnectionState State { get; init; }

    /// <summary>Process ID del worker (null se non avviato).</summary>
    public int? ProcessId { get; init; }

    /// <summary>True se ExchangeOnlineManagement module è disponibile.</summary>
    public bool IsModuleAvailable { get; init; }

    /// <summary>Versione PowerShell del worker.</summary>
    public string? PowerShellVersion { get; init; }

    /// <summary>Versione contratti del worker.</summary>
    public string? ContractsVersion { get; init; }

    /// <summary>Versione del worker.</summary>
    public string? WorkerVersion { get; init; }

    /// <summary>Numero di restart effettuati.</summary>
    public int RestartCount { get; init; }

    /// <summary>Timestamp ultimo heartbeat ricevuto.</summary>
    public DateTime? LastHeartbeat { get; init; }

    /// <summary>Numero di heartbeat mancati consecutivi.</summary>
    public int MissedHeartbeatCount { get; init; }

    /// <summary>Ultimo errore rilevato.</summary>
    public string? LastError { get; init; }
}

/// <summary>
/// Supervisore del processo worker PowerShell.
/// Gestisce avvio, monitoraggio heartbeat con grace period, restart automatico e kill.
/// </summary>
public class WorkerSupervisor : IAsyncDisposable
{
    private readonly WorkerSupervisorOptions _options;
    private readonly IpcClient _ipcClient;

    private Process? _workerProcess;
    private WorkerConnectionState _state = WorkerConnectionState.NotStarted;
    private int _restartCount;
    private DateTime? _lastHeartbeat;
    private int _missedHeartbeatCount;
    private string? _lastError;

    private CancellationTokenSource? _heartbeatCts;
    private Task? _heartbeatTask;
    private Task? _monitorTask;
    private long _heartbeatSequence;

    private HandshakeResponse? _lastHandshake;
    private readonly object _stateLock = new();
    private volatile bool _isDisposing;

    /// <summary>
    /// Evento di cambio stato del worker.
    /// </summary>
    public event EventHandler<WorkerConnectionState>? StateChanged;

    /// <summary>
    /// Evento ricevuto dal worker via IPC.
    /// </summary>
    public event EventHandler<EventEnvelope>? EventReceived;

    /// <summary>
    /// Stato corrente della connessione al worker.
    /// </summary>
    public WorkerConnectionState State => _state;

    /// <summary>
    /// Client IPC sottostante.
    /// </summary>
    public IpcClient IpcClient => _ipcClient;

    /// <summary>
    /// Crea un nuovo supervisor.
    /// </summary>
    /// <param name="options">Opzioni di configurazione (null = default).</param>
    public WorkerSupervisor(WorkerSupervisorOptions? options = null)
    {
        _options = options ?? new WorkerSupervisorOptions();
        _ipcClient = new IpcClient();
        _ipcClient.ConnectionStateChanged += OnIpcConnectionStateChanged;
        _ipcClient.EventReceived += OnEventReceived;
    }

    /// <summary>
    /// Avvia il processo worker e stabilisce connessione IPC.
    /// </summary>
    /// <param name="cancellationToken">Token di cancellazione.</param>
    /// <returns>True se avviato e connesso con successo.</returns>
    public async Task<bool> StartAsync(CancellationToken cancellationToken = default)
    {
        if (_isDisposing)
        {
            return false;
        }

        lock (_stateLock)
        {
            if (_state == WorkerConnectionState.Connected || _state == WorkerConnectionState.Starting)
            {
                return _state == WorkerConnectionState.Connected;
            }
        }

        SetState(WorkerConnectionState.Starting);

        try
        {
            // Trova path worker
            var workerPath = FindWorkerPath();
            if (string.IsNullOrEmpty(workerPath) || !File.Exists(workerPath))
            {
                _lastError = $"Worker executable not found: {_options.WorkerPath}";
                Console.WriteLine($"[Supervisor] {_lastError}");
                SetState(WorkerConnectionState.Crashed);
                return false;
            }

            Console.WriteLine($"[Supervisor] Starting worker: {workerPath}");

            // Avvia processo con capacità di interazione per l'autenticazione
            var startInfo = new ProcessStartInfo
            {
                FileName = workerPath,
                UseShellExecute = false,
                CreateNoWindow = true, // Evita la console del worker
                WindowStyle = ProcessWindowStyle.Hidden, // Nasconde la finestra se presente
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            if (!string.IsNullOrWhiteSpace(_options.ExchangeEnvironmentName))
            {
                startInfo.Environment["EXCHANGEADMIN_EXO_ENV"] = _options.ExchangeEnvironmentName;
            }

            _workerProcess = Process.Start(startInfo);
            if (_workerProcess == null)
            {
                _lastError = "Failed to start worker process";
                Console.WriteLine($"[Supervisor] {_lastError}");
                SetState(WorkerConnectionState.Crashed);
                return false;
            }

            Console.WriteLine($"[Supervisor] Worker started, PID: {_workerProcess.Id}");

            // Capture output per debug
            _workerProcess.OutputDataReceived += (s, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    Console.WriteLine($"[Worker OUT] {e.Data}");
                }
            };
            _workerProcess.ErrorDataReceived += (s, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    Console.WriteLine($"[Worker ERR] {e.Data}");
                }
            };
            _workerProcess.BeginOutputReadLine();
            _workerProcess.BeginErrorReadLine();

            SetState(WorkerConnectionState.WaitingForHandshake);

            // Attendi che la pipe sia pronta
            await Task.Delay(500, cancellationToken).ConfigureAwait(false);

            // Connetti via IPC con retry
            using var timeoutCts = new CancellationTokenSource(_options.StartupTimeoutMs);
            using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token);

            var retryCount = 0;
            const int maxRetries = 10;

            while (retryCount < maxRetries)
            {
                try
                {
                    Console.WriteLine($"[Supervisor] Attempting IPC connection (attempt {retryCount + 1}/{maxRetries})...");
                    _lastHandshake = await _ipcClient.ConnectAsync(linkedCts.Token).ConfigureAwait(false);
                    Console.WriteLine($"[Supervisor] IPC connection successful!");
                    break;
                }
                catch (TimeoutException ex) when (retryCount < maxRetries - 1)
                {
                    retryCount++;
                    Console.WriteLine($"[Supervisor] Connection timeout, retry {retryCount}/{maxRetries}: {ex.Message}");
                    await Task.Delay(500, linkedCts.Token).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[Supervisor] IPC connection failed with exception: {ex.GetType().Name}: {ex.Message}");
                    throw;
                }
            }

            if (_lastHandshake == null)
            {
                _lastError = "Failed to connect to worker after startup";
                Console.WriteLine($"[Supervisor] {_lastError}");
                SetState(WorkerConnectionState.Crashed);
                return false;
            }

            Console.WriteLine($"[Supervisor] Connected. Module available: {_lastHandshake.IsModuleAvailable}, PS version: {_lastHandshake.PowerShellVersion}");

            SetState(WorkerConnectionState.Connected);
            _restartCount = 0;
            _missedHeartbeatCount = 0;
            _lastHeartbeat = DateTime.UtcNow;

            // Avvia heartbeat e monitor
            _heartbeatCts = new CancellationTokenSource();
            _heartbeatTask = HeartbeatLoopAsync(_heartbeatCts.Token);
            _monitorTask = MonitorProcessAsync(_heartbeatCts.Token);

            return true;
        }
        catch (OperationCanceledException ex)
        {
            _lastError = "Startup cancelled";
            Console.WriteLine($"[Supervisor] Startup cancelled: {ex.Message}");
            KillWorker(); // Ensure worker is terminated
            SetState(WorkerConnectionState.Stopped);
            return false;
        }
        catch (Exception ex)
        {
            _lastError = ex.Message;
            Console.WriteLine($"[Supervisor] Start failed: {ex.GetType().Name}: {ex.Message}");
            Console.WriteLine($"[Supervisor] Stack trace: {ex.StackTrace}");
            KillWorker(); // Ensure worker is terminated
            SetState(WorkerConnectionState.Crashed);
            return false;
        }
    }

    /// <summary>
    /// Ferma il worker in modo pulito.
    /// </summary>
    public async Task StopAsync()
    {
        Console.WriteLine("[Supervisor] Stopping worker...");

        _heartbeatCts?.Cancel();

        try
        {
            await _ipcClient.DisconnectAsync().ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Supervisor] Disconnect error (ignored): {ex.Message}");
        }

        if (_workerProcess != null && !_workerProcess.HasExited)
        {
            try
            {
                Console.WriteLine($"[Supervisor] Killing worker process {_workerProcess.Id}");
                _workerProcess.Kill(entireProcessTree: true);
                using var exitCts = new CancellationTokenSource(5000);
                await _workerProcess.WaitForExitAsync(exitCts.Token).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Supervisor] Kill error (ignored): {ex.Message}");
            }
        }

        _workerProcess?.Dispose();
        _workerProcess = null;

        SetState(WorkerConnectionState.Stopped);
        Console.WriteLine("[Supervisor] Worker stopped");
    }

    /// <summary>
    /// Riavvia il worker.
    /// </summary>
    /// <param name="cancellationToken">Token di cancellazione.</param>
    /// <returns>True se riavvio riuscito.</returns>
    public async Task<bool> RestartAsync(CancellationToken cancellationToken = default)
    {
        Console.WriteLine($"[Supervisor] Restarting worker (attempt {_restartCount + 1}/{_options.MaxRestartAttempts})");

        SetState(WorkerConnectionState.Restarting);
        _restartCount++;

        await StopAsync().ConfigureAwait(false);
        await Task.Delay(_options.RestartCooldownMs, cancellationToken).ConfigureAwait(false);

        return await StartAsync(cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Kill forzato del worker senza cleanup.
    /// </summary>
    public void KillWorker()
    {
        Console.WriteLine("[Supervisor] Force killing worker");

        _heartbeatCts?.Cancel();

        if (_workerProcess != null && !_workerProcess.HasExited)
        {
            try
            {
                _workerProcess.Kill(entireProcessTree: true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Supervisor] Kill error (ignored): {ex.Message}");
            }
        }

        _workerProcess?.Dispose();
        _workerProcess = null;

        SetState(WorkerConnectionState.Stopped);
    }

    /// <summary>
    /// Ottiene informazioni sullo stato corrente del worker.
    /// </summary>
    public WorkerStatus GetStatus()
    {
        return new WorkerStatus
        {
            State = _state,
            ProcessId = _workerProcess?.Id,
            IsModuleAvailable = _lastHandshake?.IsModuleAvailable ?? false,
            PowerShellVersion = _lastHandshake?.PowerShellVersion,
            ContractsVersion = _lastHandshake?.ContractsVersion,
            WorkerVersion = _lastHandshake?.WorkerVersion,
            RestartCount = _restartCount,
            LastHeartbeat = _lastHeartbeat,
            MissedHeartbeatCount = _missedHeartbeatCount,
            LastError = _lastError
        };
    }

    private async Task HeartbeatLoopAsync(CancellationToken cancellationToken)
    {
        Console.WriteLine("[Supervisor] Heartbeat loop started");

        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                await Task.Delay(_options.HeartbeatIntervalMs, cancellationToken).ConfigureAwait(false);

                if (_state != WorkerConnectionState.Connected)
                {
                    continue;
                }

                _heartbeatSequence++;
                await _ipcClient.SendHeartbeatAsync(_heartbeatSequence, cancellationToken).ConfigureAwait(false);

                // Se riceviamo risposta (gestita nel response loop), resettiamo il contatore
                // Per ora assumiamo che l'invio riuscito sia sufficiente
                _lastHeartbeat = DateTime.UtcNow;
                _missedHeartbeatCount = 0;
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Supervisor] Heartbeat send error: {ex.Message}");
                _missedHeartbeatCount++;
            }
        }

        Console.WriteLine("[Supervisor] Heartbeat loop terminated");
    }

    private async Task MonitorProcessAsync(CancellationToken cancellationToken)
    {
        Console.WriteLine("[Supervisor] Monitor loop started");

        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                await Task.Delay(1000, cancellationToken).ConfigureAwait(false);

                if (_workerProcess == null)
                {
                    continue;
                }

                // Verifica processo ancora in esecuzione
                if (_workerProcess.HasExited)
                {
                    var exitCode = _workerProcess.ExitCode;
                    _lastError = $"Worker process exited unexpectedly with code {exitCode}";
                    Console.WriteLine($"[Supervisor] {_lastError}");

                    if (_state == WorkerConnectionState.Connected)
                    {
                        SetState(WorkerConnectionState.Crashed);

                        // Auto-restart se non superato limite
                        if (_restartCount < _options.MaxRestartAttempts)
                        {
                            await RestartAsync(cancellationToken).ConfigureAwait(false);
                        }
                        else
                        {
                            Console.WriteLine("[Supervisor] Max restart attempts reached, giving up");
                        }
                    }
                    continue;
                }

                // Verifica heartbeat timeout con grace period
                if (_state == WorkerConnectionState.Connected && _lastHeartbeat.HasValue)
                {
                    var timeSinceLastHeartbeat = DateTime.UtcNow - _lastHeartbeat.Value;
                    var totalTimeoutMs = _options.HeartbeatTimeoutMs + _options.HeartbeatGracePeriodMs;

                    // Primo livello: heartbeat timeout semplice
                    if (timeSinceLastHeartbeat.TotalMilliseconds > _options.HeartbeatTimeoutMs)
                    {
                        _missedHeartbeatCount++;
                        Console.WriteLine($"[Supervisor] Heartbeat missed ({_missedHeartbeatCount}/{_options.HeartbeatMissedThreshold})");

                        // Secondo livello: superata soglia di heartbeat mancati O grace period
                        if (_missedHeartbeatCount >= _options.HeartbeatMissedThreshold ||
                            timeSinceLastHeartbeat.TotalMilliseconds > totalTimeoutMs)
                        {
                            _lastError = $"Heartbeat timeout - worker unresponsive (missed {_missedHeartbeatCount}, last seen {timeSinceLastHeartbeat.TotalSeconds:F1}s ago)";
                            Console.WriteLine($"[Supervisor] {_lastError}");

                            SetState(WorkerConnectionState.Unresponsive);

                            // Kill e restart
                            KillWorker();

                            if (_restartCount < _options.MaxRestartAttempts)
                            {
                                await Task.Delay(_options.RestartCooldownMs, cancellationToken).ConfigureAwait(false);
                                await StartAsync(cancellationToken).ConfigureAwait(false);
                            }
                            else
                            {
                                Console.WriteLine("[Supervisor] Max restart attempts reached, giving up");
                            }
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Supervisor] Monitor error: {ex.Message}");
            }
        }

        Console.WriteLine("[Supervisor] Monitor loop terminated");
    }

    private void SetState(WorkerConnectionState newState)
    {
        lock (_stateLock)
        {
            if (_state != newState)
            {
                Console.WriteLine($"[Supervisor] State: {_state} -> {newState}");
                _state = newState;
                StateChanged?.Invoke(this, newState);
            }
        }
    }

    private void OnIpcConnectionStateChanged(object? sender, WorkerConnectionState state)
    {
        if (state == WorkerConnectionState.Crashed && _state == WorkerConnectionState.Connected)
        {
            Console.WriteLine("[Supervisor] IPC connection crashed");
            SetState(WorkerConnectionState.Crashed);
        }
    }

    private void OnEventReceived(object? sender, EventEnvelope evt)
    {
        EventReceived?.Invoke(this, evt);
    }

    private string FindWorkerPath()
    {
        // Cerca in vari percorsi possibili
        var candidates = new[]
        {
            _options.WorkerPath,
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, _options.WorkerPath),
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExchangeAdmin.Worker.exe"),
            Path.Combine(Directory.GetCurrentDirectory(), _options.WorkerPath),
            Path.Combine(Directory.GetCurrentDirectory(), "ExchangeAdmin.Worker.exe")
        };

        foreach (var candidate in candidates)
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return _options.WorkerPath;
    }

    /// <summary>
    /// Rilascia le risorse asincronamente.
    /// </summary>
    public async ValueTask DisposeAsync()
    {
        if (_isDisposing)
        {
            return;
        }

        _isDisposing = true;

        await StopAsync().ConfigureAwait(false);

        // Dispose del CancellationTokenSource
        try
        {
            _heartbeatCts?.Cancel();
            _heartbeatCts?.Dispose();
            _heartbeatCts = null;
        }
        catch
        {
            // Ignora errori di dispose
        }

        await _ipcClient.DisposeAsync().ConfigureAwait(false);
    }
}
