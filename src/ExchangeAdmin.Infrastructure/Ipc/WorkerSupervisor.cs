using System.Diagnostics;
using ExchangeAdmin.Contracts;
using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Infrastructure.Ipc;

             
                                                           
              
public class WorkerSupervisorOptions
{
                 
                                                                        
                  
    public string WorkerPath { get; set; } = "ExchangeAdmin.Worker.exe";

                 
                                                                 
                  
    public int MaxRestartAttempts { get; set; } = 3;

                 
                                                               
                  
    public int RestartCooldownMs { get; set; } = 2000;

                 
                                                     
                  
    public int StartupTimeoutMs { get; set; } = 30000;

                 
                                           
                  
    public int HeartbeatIntervalMs { get; set; } = IpcConstants.HeartbeatIntervalMs;

                 
                                                                                                   
                  
    public int HeartbeatTimeoutMs { get; set; } = IpcConstants.HeartbeatTimeoutMs;

                 
                                                                                  
                                  
                  
    public int HeartbeatGracePeriodMs { get; set; } = IpcConstants.HeartbeatGracePeriodMs;

                 
                                                              
                  
    public int HeartbeatMissedThreshold { get; set; } = IpcConstants.HeartbeatMissedThreshold;

                 
                                                                        
                                                          
                  
    public string? ExchangeEnvironmentName { get; set; }
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
}

             
                                               
                                                                                       
              
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

                 
                                          
                  
    public event EventHandler<WorkerConnectionState>? StateChanged;

                 
                                           
                  
    public event EventHandler<EventEnvelope>? EventReceived;

                 
                                                   
                  
    public WorkerConnectionState State => _state;

                 
                               
                  
    public IpcClient IpcClient => _ipcClient;

                 
                                 
                  
                                                                                 
    public WorkerSupervisor(WorkerSupervisorOptions? options = null)
    {
        _options = options ?? new WorkerSupervisorOptions();
        _ipcClient = new IpcClient();
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
                Console.WriteLine($"[Supervisor] {_lastError}");
                SetState(WorkerConnectionState.Crashed);
                return false;
            }

            Console.WriteLine($"[Supervisor] Starting worker: {workerPath}");

            var startInfo = new ProcessStartInfo
            {
                FileName = workerPath,
                UseShellExecute = true,
                CreateNoWindow = false,
                WindowStyle = ProcessWindowStyle.Normal
            };

            if (!string.IsNullOrWhiteSpace(_options.ExchangeEnvironmentName))
            {
                startInfo.EnvironmentVariables["EXCHANGEADMIN_EXO_ENV"] = _options.ExchangeEnvironmentName;
                startInfo.UseShellExecute = false;
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

            SetState(WorkerConnectionState.WaitingForHandshake);

                                             
            await Task.Delay(500, cancellationToken).ConfigureAwait(false);

                                         
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

                                        
            _heartbeatCts = new CancellationTokenSource();
            _heartbeatTask = HeartbeatLoopAsync(_heartbeatCts.Token);
            _monitorTask = MonitorProcessAsync(_heartbeatCts.Token);

            return true;
        }
        catch (OperationCanceledException ex)
        {
            _lastError = "Startup cancelled";
            Console.WriteLine($"[Supervisor] Startup cancelled: {ex.Message}");
            KillWorker();                               
            SetState(WorkerConnectionState.Stopped);
            return false;
        }
        catch (Exception ex)
        {
            _lastError = ex.Message;
            Console.WriteLine($"[Supervisor] Start failed: {ex.GetType().Name}: {ex.Message}");
            Console.WriteLine($"[Supervisor] Stack trace: {ex.StackTrace}");
            KillWorker();                               
            SetState(WorkerConnectionState.Crashed);
            return false;
        }
    }

                 
                                       
                  
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

                 
                          
                  
                                                                       
                                                    
    public async Task<bool> RestartAsync(CancellationToken cancellationToken = default)
    {
        Console.WriteLine($"[Supervisor] Restarting worker (attempt {_restartCount + 1}/{_options.MaxRestartAttempts})");

        SetState(WorkerConnectionState.Restarting);
        _restartCount++;

        await StopAsync().ConfigureAwait(false);
        await Task.Delay(_options.RestartCooldownMs, cancellationToken).ConfigureAwait(false);

        return await StartAsync(cancellationToken).ConfigureAwait(false);
    }

                 
                                              
                  
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

                                                         
                if (_workerProcess.HasExited)
                {
                    var exitCode = _workerProcess.ExitCode;
                    _lastError = $"Worker process exited unexpectedly with code {exitCode}";
                    Console.WriteLine($"[Supervisor] {_lastError}");

                    if (_state == WorkerConnectionState.Connected)
                    {
                        SetState(WorkerConnectionState.Crashed);

                                                              
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

                                                              
                if (_state == WorkerConnectionState.Connected && _lastHeartbeat.HasValue)
                {
                    var timeSinceLastHeartbeat = DateTime.UtcNow - _lastHeartbeat.Value;
                    var totalTimeoutMs = _options.HeartbeatTimeoutMs + _options.HeartbeatGracePeriodMs;

                                                                
                    if (timeSinceLastHeartbeat.TotalMilliseconds > _options.HeartbeatTimeoutMs)
                    {
                        _missedHeartbeatCount++;
                        Console.WriteLine($"[Supervisor] Heartbeat missed ({_missedHeartbeatCount}/{_options.HeartbeatMissedThreshold})");

                                                                                               
                        if (_missedHeartbeatCount >= _options.HeartbeatMissedThreshold ||
                            timeSinceLastHeartbeat.TotalMilliseconds > totalTimeoutMs)
                        {
                            _lastError = $"Heartbeat timeout - worker unresponsive (missed {_missedHeartbeatCount}, last seen {timeSinceLastHeartbeat.TotalSeconds:F1}s ago)";
                            Console.WriteLine($"[Supervisor] {_lastError}");

                            SetState(WorkerConnectionState.Unresponsive);

                                             
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
