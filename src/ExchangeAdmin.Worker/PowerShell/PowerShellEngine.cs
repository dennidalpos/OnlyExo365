using System.Diagnostics;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using ExchangeAdmin.Contracts;

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
    internal const string ExchangeOnlineModuleName = "ExchangeOnlineManagement";
    internal const int MaxConsecutiveFailuresBeforeReset = 3;

    private Runspace? _runspace;
    private bool _isModuleAvailable;
    private string? _powerShellVersion;
    private bool _isInitialized;
    private bool _isConnected;
    private bool _isGraphConnected;
    private bool _isComplianceConnected;
    private IReadOnlyList<string> _connectedGraphScopes = Array.Empty<string>();
    private readonly SemaphoreSlim _executionLock = new(1, 1);
    private readonly object _stateLock = new();
    private readonly ExchangeOnlineConfiguration _exchangeConfiguration;
    private volatile bool _isDisposing;
    private int _consecutiveFailures;
    private readonly RunspaceLifecycleManager _runspaceLifecycleManager;
    private readonly RunspaceRecoveryService _runspaceRecoveryService;
    private readonly PowerShellExecutionPipeline _executionPipeline;
    private readonly ExchangeConnectionSession _connectionSession;

    public bool IsInitialized => _isInitialized;

    public bool IsConnected => _isConnected;

    public bool IsGraphConnected => _isGraphConnected;

    public bool IsComplianceConnected => _isComplianceConnected;

    public IReadOnlyList<string> ConnectedGraphScopes => _connectedGraphScopes;

    public bool IsModuleAvailable => _isModuleAvailable;

    public string? PowerShellVersion => _powerShellVersion;

    public PowerShellEngine(ExchangeOnlineConfiguration? exchangeConfiguration = null)
    {
        _exchangeConfiguration = exchangeConfiguration?.Clone() ?? ExchangeOnlineConfiguration.CreateDefault();
        _runspaceLifecycleManager = new RunspaceLifecycleManager(this);
        _runspaceRecoveryService = new RunspaceRecoveryService(this, _runspaceLifecycleManager);
        _executionPipeline = new PowerShellExecutionPipeline(this, _runspaceRecoveryService);
        _connectionSession = new ExchangeConnectionSession(this, _executionPipeline);
    }

    internal Runspace? Runspace
    {
        get => _runspace;
        set => _runspace = value;
    }

    internal bool ModuleAvailable
    {
        get => _isModuleAvailable;
        set => _isModuleAvailable = value;
    }

    internal string? PowerShellVersionValue
    {
        get => _powerShellVersion;
        set => _powerShellVersion = value;
    }

    internal bool Initialized
    {
        get => _isInitialized;
        set => _isInitialized = value;
    }

    internal bool Connected
    {
        get => _isConnected;
        set => _isConnected = value;
    }

    internal bool GraphConnected
    {
        get => _isGraphConnected;
        set => _isGraphConnected = value;
    }

    internal bool ComplianceConnected
    {
        get => _isComplianceConnected;
        set => _isComplianceConnected = value;
    }

    internal IReadOnlyList<string> ConnectedGraphScopesValue
    {
        get => _connectedGraphScopes;
        set => _connectedGraphScopes = value;
    }

    internal SemaphoreSlim ExecutionLock => _executionLock;

    internal object StateLock => _stateLock;

    internal ExchangeOnlineConfiguration ExchangeConfiguration => _exchangeConfiguration;

    internal bool IsDisposingRequested => _isDisposing;

    internal int ConsecutiveFailures
    {
        get => _consecutiveFailures;
        set => _consecutiveFailures = value;
    }

    public Task<PowerShellInitResult> InitializeAsync()
    {
        return _runspaceLifecycleManager.InitializeAsync();
    }

    public Task<PowerShellResult> ExecuteAsync(
        string script,
        Dictionary<string, object>? parameters = null,
        Action<string, string>? onVerbose = null,
        Action<string, string>? onWarning = null,
        Action<ErrorRecord>? onError = null,
        Action<PSObject>? onOutput = null,
        CancellationToken cancellationToken = default)
    {
        return _executionPipeline.ExecuteAsync(
            script,
            parameters,
            onVerbose,
            onWarning,
            onError,
            onOutput,
            cancellationToken);
    }

    public Task<PowerShellResult> ConnectExchangeAsync(
        Action<string, string>? onVerbose = null,
        CancellationToken cancellationToken = default)
    {
        return _connectionSession.ConnectExchangeAsync(onVerbose, cancellationToken);
    }

    public Task<PowerShellResult> DisconnectExchangeAsync(CancellationToken cancellationToken = default)
    {
        return _connectionSession.DisconnectExchangeAsync(cancellationToken);
    }

    public Task<PowerShellResult> ConnectMicrosoftGraphAsync(
        bool ignoreAutoConnectConfiguration = false,
        IEnumerable<string>? delegatedScopes = null,
        Action<string, string>? onVerbose = null,
        CancellationToken cancellationToken = default)
    {
        return _connectionSession.ConnectMicrosoftGraphAsync(
            ignoreAutoConnectConfiguration,
            delegatedScopes,
            onVerbose,
            cancellationToken);
    }

    public Task<PowerShellResult> ConnectComplianceAsync(
        Action<string, string>? onVerbose = null,
        CancellationToken cancellationToken = default)
    {
        return _connectionSession.ConnectComplianceAsync(onVerbose, cancellationToken);
    }

    public Task<(bool IsConnected, string? UserPrincipalName, string? Organization, bool IsGraphConnected, bool IsComplianceConnected)> GetConnectionStatusAsync(
        CancellationToken cancellationToken = default)
    {
        return _connectionSession.GetConnectionStatusAsync(cancellationToken);
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

        _runspaceLifecycleManager.DisposeRunspace();

        Debug.WriteLine("[PowerShellEngine] Disposed");
    }
}
