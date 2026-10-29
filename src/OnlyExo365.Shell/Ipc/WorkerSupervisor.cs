using System.IO;
using System.Diagnostics;
using OnlyExo365.Contracts.Diagnostics;
using OnlyExo365.Contracts;
using OnlyExo365.Contracts.Messages;

namespace OnlyExo365.Shell.Ipc;

             
                                                           
              
public class WorkerSupervisorOptions
{
                 
                                                                        
                  
    public string WorkerPath { get; set; } = "OnlyExo365.Worker.exe";

                 
                                                                 
                  
    public int MaxRestartAttempts { get; set; } = 3;

                 
                                                               
                  
    public int RestartCooldownMs { get; set; } = 2000;

                 
                                                     
                  
    public int StartupTimeoutMs { get; set; } = 90000;

                 
                                           
                  
    public int HeartbeatIntervalMs { get; set; } = IpcConstants.HeartbeatIntervalMs;

                 
                                                                                                   
                  
    public int HeartbeatTimeoutMs { get; set; } = IpcConstants.HeartbeatTimeoutMs;

                 
                                                                                  
                                  
                  
    public int HeartbeatGracePeriodMs { get; set; } = IpcConstants.HeartbeatGracePeriodMs;

                 
                                                              
                  
    public int HeartbeatMissedThreshold { get; set; } = IpcConstants.HeartbeatMissedThreshold;

                 
                                                                        
                                                          
                  
    public ExchangeOnlineConfiguration ExchangeConfiguration { get; set; } = ExchangeOnlineConfiguration.CreateDefault();

    public string? ExchangeEnvironmentName
    {
        get => ExchangeConfiguration.ExchangeEnvironmentName;
        set => ExchangeConfiguration.ExchangeEnvironmentName = string.IsNullOrWhiteSpace(value) ? "O365Default" : value.Trim();
    }
}

             
                                                 
              
public class WorkerStatus
{
                                                      
    public WorkerConnectionState State { get; init; }

                                                                       
    public int? ProcessId { get; init; }

                                                                                  
    public bool IsModuleAvailable { get; init; }

                                                          
    public string? PowerShellVersion { get; init; }

                                                         
    public string? ContractsVersion { get; init; }

                                               
    public string? WorkerVersion { get; init; }

                                                        
    public int RestartCount { get; init; }

                                                               
    public DateTime? LastHeartbeat { get; init; }

                                                                   
    public int MissedHeartbeatCount { get; init; }

                                                  
    public string? LastError { get; init; }

    public bool IsConsoleVisible { get; init; }
}

             
                                               
                                                                                       
              
public class WorkerSupervisor : IAsyncDisposable
{
    private readonly WorkerSupervisorOptions _options;
    private readonly IpcClient _ipcClient;
    private readonly IpcSessionContext _ipcSessionContext;
    private readonly string _ipcSessionToken;
    private readonly PersistentLogWriter _persistentLogWriter = new("supervisor");

    private Process? _workerProcess;
    private WorkerConnectionState _state = WorkerConnectionState.NotStarted;
    private int _restartCount;
    private DateTime? _lastHeartbeat;
    private int _missedHeartbeatCount;
    private string? _lastError;
    private bool _isConsoleVisible;

    private CancellationTokenSource? _heartbeatCts;
    private Task? _heartbeatTask;
    private Task? _monitorTask;
    private long _heartbeatSequence;
    private readonly object _stopLock = new();
    private Task? _stopTask;

    private HandshakeResponse? _lastHandshake;
    private readonly object _stateLock = new();
    private volatile bool _isDisposing;
    private volatile bool _isStopping;

                 
                                          
                  
    public event EventHandler<WorkerConnectionState>? StateChanged;

                 
                                           
                  
    public event EventHandler<EventEnvelope>? EventReceived;

                 
                                                   
                  
    public WorkerConnectionState State => _state;

                 
                               
                  
    public IpcClient IpcClient => _ipcClient;

                 
                                 
                  
                                                                                 
    public WorkerSupervisor(WorkerSupervisorOptions? options = null)
    {
        _options = options ?? new WorkerSupervisorOptions();
        _ipcSessionContext = IpcSessionContext.CreateForCurrentProcess();
        _ipcSessionToken = Guid.NewGuid().ToString("N");
        _ipcClient = new IpcClient(_ipcSessionContext, _ipcSessionToken);
        _ipcClient.ConnectionStateChanged += OnIpcConnectionStateChanged;
        _ipcClient.EventReceived += OnEventReceived;
        _ipcClient.HeartbeatReceived += OnHeartbeatReceived;
    }

                 
                                                              
                  
                                                                       
                                                                   
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
                                
            var workerPath = FindWorkerPath();
            if (string.IsNullOrEmpty(workerPath) || !File.Exists(workerPath))
            {
                _lastError = $"Worker executable not found: {_options.WorkerPath}";
                LogError(_lastError);
                SetState(WorkerConnectionState.Crashed);
                return false;
            }

            LogInformation($"Starting worker: {workerPath}");

            var startInfo = new ProcessStartInfo
            {
                FileName = workerPath,
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            startInfo.EnvironmentVariables[IpcConstants.SessionTokenEnvironmentVariable] = _ipcSessionToken;

            foreach (var validationError in _options.ExchangeConfiguration.Validate())
            {
                LogWarning($"Exchange configuration warning: {validationError}");
            }

            _options.ExchangeConfiguration.ApplyEnvironmentVariables(startInfo.EnvironmentVariables);

            _workerProcess = Process.Start(startInfo);
            if (_workerProcess == null)
            {
                _lastError = "Failed to start worker process";
                LogError(_lastError);
                SetState(WorkerConnectionState.Crashed);
                return false;
            }

            LogInformation($"Worker started, PID: {_workerProcess.Id}");
            _isConsoleVisible = false;

            SetState(WorkerConnectionState.WaitingForHandshake);

                                             
            await Task.Delay(500, cancellationToken).ConfigureAwait(false);

                                         
            _lastHandshake = null;

            using var timeoutCts = new CancellationTokenSource(_options.StartupTimeoutMs);
            using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token);

            var retryCount = 0;
            const int maxRetries = 10;

            while (retryCount < maxRetries)
            {
                try
                {
                    LogDebug($"Attempting IPC connection (attempt {retryCount + 1}/{maxRetries})...");
                    _lastHandshake = await _ipcClient.ConnectAsync(linkedCts.Token).ConfigureAwait(false);
                    LogInformation("IPC connection successful");
                    break;
                }
                catch (TimeoutException ex) when (retryCount < maxRetries - 1)
                {
                    retryCount++;
                    LogWarning($"Connection timeout, retry {retryCount}/{maxRetries}: {ex.Message}");
                    await Task.Delay(500, linkedCts.Token).ConfigureAwait(false);
                }
                catch (OperationCanceledException) when (timeoutCts.IsCancellationRequested)
                {
                    _lastError = $"IPC startup timeout after {_options.StartupTimeoutMs} ms";
                    LogError(_lastError);
                    break;
                }
                catch (Exception ex)
                {
                    LogError($"IPC connection failed with exception: {ex.GetType().Name}: {ex.Message}");
                    throw;
                }
            }

            if (_lastHandshake == null && _workerProcess is { HasExited: false })
            {
                var fallbackTimeoutMs = Math.Max(30000, _options.StartupTimeoutMs / 2);
                LogWarning($"Worker process is still running. Retrying IPC handshake for {fallbackTimeoutMs} ms before failing startup...");

                using var fallbackCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                fallbackCts.CancelAfter(fallbackTimeoutMs);

                try
                {
                    _lastHandshake = await _ipcClient.ConnectAsync(fallbackCts.Token).ConfigureAwait(false);
                    LogInformation("IPC connection successful on fallback attempt");
                }
                catch (Exception ex)
                {
                    LogError($"Fallback IPC connection failed: {ex.GetType().Name}: {ex.Message}");
                }
            }

            if (_lastHandshake == null)
            {
                _lastError = "Failed to connect to worker after startup";
                LogError(_lastError);
                KillWorker();
                SetState(WorkerConnectionState.Crashed);
                return false;
            }

            LogInformation($"Connected. Module available: {_lastHandshake.IsModuleAvailable}, PS version: {_lastHandshake.PowerShellVersion}");

            SetState(WorkerConnectionState.Connected);
            _restartCount = 0;
            _missedHeartbeatCount = 0;
            _lastHeartbeat = DateTime.UtcNow;

                                        
            _heartbeatCts = new CancellationTokenSource();
            _heartbeatTask = HeartbeatLoopAsync(_heartbeatCts.Token);
            _monitorTask = MonitorProcessAsync(_heartbeatCts.Token);

            return true;
        }
        catch (OperationCanceledException ex)
        {
            _lastError = "Startup cancelled";
            LogWarning($"Startup cancelled: {ex.Message}");
            KillWorker();                               
            SetState(WorkerConnectionState.Stopped);
            return false;
        }
        catch (Exception ex)
        {
            _lastError = ex.Message;
            LogError($"Start failed: {ex.GetType().Name}: {ex.Message}");
            LogDebug($"Stack trace: {ex.StackTrace}");
            KillWorker();                               
            SetState(WorkerConnectionState.Crashed);
            return false;
        }
    }

                 
                                       
                  
    public Task StopAsync()
    {
        lock (_stopLock)
        {
            _stopTask ??= StopCoreAsync();
            return _stopTask;
        }
    }

    private async Task StopCoreAsync()
    {
        _isStopping = true;
        try
        {
            LogInformation("Stopping worker...");

            _heartbeatCts?.Cancel();

            try
            {
                await _ipcClient.DisconnectAsync().ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                LogWarning($"Disconnect error (ignored): {ex.Message}");
            }

            if (_workerProcess != null && !_workerProcess.HasExited)
            {
                try
                {
                    LogWarning($"Killing worker process {_workerProcess.Id}");
                    _workerProcess.Kill(entireProcessTree: true);
                    using var exitCts = new CancellationTokenSource(5000);
                    await _workerProcess.WaitForExitAsync(exitCts.Token).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    LogWarning($"Kill error (ignored): {ex.Message}");
                }
            }

            var backgroundTasks = new[] { _heartbeatTask, _monitorTask }
                .Where(static task => task is not null)
                .Cast<Task>()
                .ToArray();

            if (backgroundTasks.Length > 0)
            {
                try
                {
                    await Task.WhenAll(backgroundTasks).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    LogWarning($"Background worker loop shutdown error (ignored): {ex.Message}");
                }
            }

            _workerProcess?.Dispose();
            _workerProcess = null;
            _isConsoleVisible = false;
            _heartbeatTask = null;
            _monitorTask = null;

            SetState(WorkerConnectionState.Stopped);
            LogInformation("Worker stopped");
        }
        finally
        {
            lock (_stopLock)
            {
                _stopTask = null;
            }

            _isStopping = false;
        }
    }

                 
                          
                  
                                                                       
                                                    
    public async Task<bool> RestartAsync(CancellationToken cancellationToken = default)
    {
        LogWarning($"Restarting worker (attempt {_restartCount + 1}/{_options.MaxRestartAttempts})");

        SetState(WorkerConnectionState.Restarting);
        _restartCount++;

        await StopAsync().ConfigureAwait(false);
        await Task.Delay(_options.RestartCooldownMs, cancellationToken).ConfigureAwait(false);

        return await StartAsync(cancellationToken).ConfigureAwait(false);
    }

    internal static async Task<bool> RestartWithinBudgetAsync(
        int restartCount,
        int maxRestartAttempts,
        Func<CancellationToken, Task<bool>> restartAction,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(restartAction);

        if (restartCount >= maxRestartAttempts)
        {
            return false;
        }

        return await restartAction(cancellationToken).ConfigureAwait(false);
    }

                 
                                              
                  
    public void KillWorker()
    {
        LogWarning("Force killing worker");

        _heartbeatCts?.Cancel();

        if (_workerProcess != null && !_workerProcess.HasExited)
        {
            try
            {
                _workerProcess.Kill(entireProcessTree: true);
            }
            catch (Exception ex)
            {
                LogWarning($"Kill error (ignored): {ex.Message}");
            }
        }

        _workerProcess?.Dispose();
        _workerProcess = null;
        _isConsoleVisible = false;

        SetState(WorkerConnectionState.Stopped);
    }

                 
                                                             
                  
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
            LastError = _lastError,
            IsConsoleVisible = _isConsoleVisible
        };
    }

    internal void SetConsoleVisibility(bool isVisible)
    {
        lock (_stateLock)
        {
            _isConsoleVisible = isVisible;
        }
    }

    private async Task HeartbeatLoopAsync(CancellationToken cancellationToken)
    {
        LogDebug("Heartbeat loop started");

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
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception ex)
            {
                LogWarning($"Heartbeat send error: {ex.Message}");
                _missedHeartbeatCount++;
            }
        }

        LogDebug("Heartbeat loop terminated");
    }

    private async Task MonitorProcessAsync(CancellationToken cancellationToken)
    {
        LogDebug("Monitor loop started");

        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                await Task.Delay(1000, cancellationToken).ConfigureAwait(false);

                if (_workerProcess == null)
                {
                    continue;
                }

                                                         
                if (_workerProcess.HasExited)
                {
                    var exitCode = _workerProcess.ExitCode;
                    _lastError = $"Worker process exited unexpectedly with code {exitCode}";
                    LogError(_lastError);

                    if (_state == WorkerConnectionState.Connected)
                    {
                        if (IsShutdownInProgress(cancellationToken))
                        {
                            LogInformation("Worker exited while shutdown was already in progress.");
                            SetState(WorkerConnectionState.Stopped);
                            continue;
                        }

                        SetState(WorkerConnectionState.Crashed);

                        if (!await RestartWithinBudgetAsync(_restartCount, _options.MaxRestartAttempts, RestartAsync, cancellationToken).ConfigureAwait(false))
                        {
                            LogError("Max restart attempts reached, giving up");
                        }
                    }
                    continue;
                }

                                                              
                if (_state == WorkerConnectionState.Connected && _lastHeartbeat.HasValue)
                {
                    var timeSinceLastHeartbeat = DateTime.UtcNow - _lastHeartbeat.Value;
                    var totalTimeoutMs = _options.HeartbeatTimeoutMs + _options.HeartbeatGracePeriodMs;

                                                                
                    if (timeSinceLastHeartbeat.TotalMilliseconds > _options.HeartbeatTimeoutMs)
                    {
                        _missedHeartbeatCount++;
                        LogWarning($"Heartbeat missed ({_missedHeartbeatCount}/{_options.HeartbeatMissedThreshold})");

                                                                                               
                        if (_missedHeartbeatCount >= _options.HeartbeatMissedThreshold ||
                            timeSinceLastHeartbeat.TotalMilliseconds > totalTimeoutMs)
                        {
                            _lastError = $"Heartbeat timeout - worker unresponsive (missed {_missedHeartbeatCount}, last seen {timeSinceLastHeartbeat.TotalSeconds:F1}s ago)";
                            LogError(_lastError);

                            SetState(WorkerConnectionState.Unresponsive);

                                             
                            KillWorker();

                            if (IsShutdownInProgress(cancellationToken))
                            {
                                break;
                            }

                            if (!await RestartWithinBudgetAsync(_restartCount, _options.MaxRestartAttempts, RestartAsync, cancellationToken).ConfigureAwait(false))
                            {
                                LogError("Max restart attempts reached, giving up");
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
                LogError($"Monitor error: {ex.Message}");
            }
        }

        LogDebug("Monitor loop terminated");
    }

    private void OnHeartbeatReceived(object? sender, HeartbeatPong pong)
    {
        if (_isDisposing)
        {
            return;
        }

        lock (_stateLock)
        {
            if (_state != WorkerConnectionState.Connected)
            {
                return;
            }

            _lastHeartbeat = DateTime.UtcNow;
            _missedHeartbeatCount = 0;
        }
    }

    private void SetState(WorkerConnectionState newState)
    {
        lock (_stateLock)
        {
            if (_state != newState)
            {
                LogInformation($"State: {_state} -> {newState}");
                _state = newState;
                StateChanged?.Invoke(this, newState);
            }
        }
    }

    private void OnIpcConnectionStateChanged(object? sender, WorkerConnectionState state)
    {
        if (state == WorkerConnectionState.Crashed && _state == WorkerConnectionState.Connected)
        {
            LogError("IPC connection crashed");
            SetState(WorkerConnectionState.Crashed);
        }
    }

    private void OnEventReceived(object? sender, EventEnvelope evt)
    {
        EventReceived?.Invoke(this, evt);
    }

    private bool IsShutdownInProgress(CancellationToken cancellationToken)
        => _isStopping || _isDisposing || cancellationToken.IsCancellationRequested;

    private string FindWorkerPath()
    {
        return ResolveWorkerPath(_options.WorkerPath, AppContext.BaseDirectory);
    }

    internal static string ResolveWorkerPath(string? configuredWorkerPath, string? baseDirectory)
    {
        var effectiveWorkerPath = string.IsNullOrWhiteSpace(configuredWorkerPath)
            ? "OnlyExo365.Worker.exe"
            : configuredWorkerPath.Trim();
        var effectiveBaseDirectory = string.IsNullOrWhiteSpace(baseDirectory)
            ? AppContext.BaseDirectory
            : Path.GetFullPath(baseDirectory);

        var candidates = new[]
        {
            Path.IsPathRooted(effectiveWorkerPath)
                ? Path.GetFullPath(effectiveWorkerPath)
                : Path.GetFullPath(Path.Combine(effectiveBaseDirectory, effectiveWorkerPath)),
            Path.GetFullPath(Path.Combine(effectiveBaseDirectory, "OnlyExo365.Worker.exe"))
        }.Distinct(StringComparer.OrdinalIgnoreCase);

        foreach (var candidate in candidates)
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return candidates.First();
    }

    private void LogInformation(string message) => WriteLog(LogLevel.Information, message);
    private void LogWarning(string message) => WriteLog(LogLevel.Warning, message);
    private void LogError(string message) => WriteLog(LogLevel.Error, message);
    private void LogDebug(string message) => WriteLog(LogLevel.Debug, message);

    private void WriteLog(LogLevel level, string message)
    {
        _persistentLogWriter.Write(level, "Supervisor", message);
        Console.WriteLine($"[Supervisor] {message}");
    }

                 
                                           
                  
    public async ValueTask DisposeAsync()
    {
        if (_isDisposing)
        {
            return;
        }

        _isDisposing = true;

        await StopAsync().ConfigureAwait(false);

                                              
        try
        {
            _heartbeatCts?.Cancel();
            _heartbeatCts?.Dispose();
            _heartbeatCts = null;
        }
        catch
        {
                                       
        }

        await _ipcClient.DisposeAsync().ConfigureAwait(false);
    }
}

