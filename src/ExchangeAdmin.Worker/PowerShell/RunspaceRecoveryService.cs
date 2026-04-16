using System.Diagnostics;
using System.Management.Automation.Runspaces;

namespace ExchangeAdmin.Worker.PowerShell;

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

    public async Task<bool> TryRecoverRunspaceAsync()
    {
        Debug.WriteLine("[PowerShellEngine] Attempting runspace recovery...");

        try
        {
            await _runspaceLifecycleManager.RecreateRunspaceAsync(_engine.ModuleAvailable).ConfigureAwait(false);

            _engine.ConsecutiveFailures = 0;
            lock (_engine.StateLock)
            {
                _engine.Connected = false;
            }

            _engine.GraphConnected = false;

            Debug.WriteLine("[PowerShellEngine] Runspace recovery successful");
            return true;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[PowerShellEngine] Runspace recovery failed: {ex.Message}");
            return false;
        }
    }
}
