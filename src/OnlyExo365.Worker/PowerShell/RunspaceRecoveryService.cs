using System.Diagnostics;
using System.Management.Automation.Runspaces;

namespace OnlyExo365.Worker.PowerShell;

internal sealed class RunspaceRecoveryService
{
    private readonly PowerShellEngine _engine;
    private readonly RunspaceLifecycleManager _runspaceLifecycleManager;

    public RunspaceRecoveryService(PowerShellEngine engine, RunspaceLifecycleManager runspaceLifecycleManager)
    {
        _engine = engine;
        _runspaceLifecycleManager = runspaceLifecycleManager;
    }

    public bool IsRunspaceUsable()
    {
        if (_engine.Runspace == null)
        {
            return false;
        }

        try
        {
            return _engine.Runspace.RunspaceStateInfo.State == RunspaceState.Opened;
        }
        catch
        {
            return false;
        }
    }

    public async Task<bool> TryRecoverRunspaceAsync(CancellationToken cancellationToken)
    {
        Debug.WriteLine("[PowerShellEngine] Attempting runspace recovery...");

        // Every connection belongs to the discarded runspace, even if recreation fails.
        lock (_engine.StateLock)
        {
            _engine.Connected = false;
            _engine.GraphConnected = false;
            _engine.ComplianceConnected = false;
            _engine.ConnectedGraphScopesValue = Array.Empty<string>();
        }

        try
        {
            await _runspaceLifecycleManager.RecreateRunspaceAsync(_engine.ModuleAvailable, cancellationToken).ConfigureAwait(false);

            _engine.ConsecutiveFailures = 0;
            Debug.WriteLine("[PowerShellEngine] Runspace recovery successful");
            return true;
        }
        catch (OperationCanceledException) { throw; }
        catch (TimeoutException) { throw; }
        catch (Exception ex)
        {
            Debug.WriteLine($"[PowerShellEngine] Runspace recovery failed: {ex.Message}");
            return false;
        }
    }
}

