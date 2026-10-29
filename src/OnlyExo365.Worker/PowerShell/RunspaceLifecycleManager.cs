using OnlyExo365.Contracts;
using System.Diagnostics;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace OnlyExo365.Worker.PowerShell;

internal sealed class RunspaceLifecycleManager
{
    private readonly PowerShellEngine _engine;

    public RunspaceLifecycleManager(PowerShellEngine engine)
    {
        _engine = engine;
    }

    public async Task<PowerShellInitResult> InitializeAsync(CancellationToken cancellationToken)
    {
        var callerCancellation = cancellationToken;
        using var budget = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        budget.CancelAfter(TimeSpan.FromSeconds(30));
        cancellationToken = budget.Token;
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            Debug.WriteLine("[PowerShellEngine] Initializing...");

            var runspace = CreateRunspace();
            _engine.Runspace = runspace;

            Debug.WriteLine($"[PowerShellEngine] Runspace opened, state: {runspace.RunspaceStateInfo.State}");

            using var ps = System.Management.Automation.PowerShell.Create();
            ps.Runspace = runspace;
            ps.AddScript("$PSVersionTable.PSVersion.ToString()");

            var versionResult = await InvokeAsync(ps, cancellationToken).ConfigureAwait(false);
            _engine.PowerShellVersionValue = versionResult.FirstOrDefault()?.ToString() ?? "Unknown";

            Debug.WriteLine($"[PowerShellEngine] PowerShell version: {_engine.PowerShellVersionValue}");

            ps.Commands.Clear();
            var approvedVersion = PowerShellModuleBootstrapPolicy.Resolve(PowerShellEngine.ExchangeOnlineModuleName)!.RequiredVersion;
            ps.AddScript($"Get-Module -ListAvailable -Name {PowerShellEngine.ExchangeOnlineModuleName} | Where-Object {{ $_.Version -eq [Version]'{approvedVersion}' }} | Select-Object -First 1");

            var moduleResult = await InvokeAsync(ps, cancellationToken).ConfigureAwait(false);
            _engine.ModuleAvailable = moduleResult.Any();

            Debug.WriteLine($"[PowerShellEngine] ExchangeOnlineManagement module available: {_engine.ModuleAvailable}");

            if (_engine.ModuleAvailable)
            {
                var packageManagementReady = await EnsurePackageManagementAvailableAsync(runspace, cancellationToken).ConfigureAwait(false);
                if (!packageManagementReady)
                {
                    Debug.WriteLine("[PowerShellEngine] PackageManagement module was not explicitly loaded. Continuing with best-effort import.");
                }

                var imported = await ImportModuleAsync(runspace, PowerShellEngine.ExchangeOnlineModuleName, cancellationToken).ConfigureAwait(false);
                _engine.ModuleAvailable = imported;
                if (!imported)
                {
                    Debug.WriteLine($"[PowerShellEngine] Warning: {PowerShellEngine.ExchangeOnlineModuleName} import reported warnings/errors");
                }
            }

            _engine.Initialized = true;
            _engine.ConsecutiveFailures = 0;

            return new PowerShellInitResult
            {
                Success = true,
                PowerShellVersion = _engine.PowerShellVersionValue,
                IsModuleAvailable = _engine.ModuleAvailable
            };
        }
        catch (OperationCanceledException ex) when (!callerCancellation.IsCancellationRequested)
        {
            throw new TimeoutException("PowerShell initialization timed out.", ex);
        }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex)
        {
            Debug.WriteLine($"[PowerShellEngine] Initialization failed: {ex.Message}");

            return new PowerShellInitResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                PowerShellVersion = _engine.PowerShellVersionValue,
                IsModuleAvailable = false
            };
        }
    }

    public async Task<bool> RecreateRunspaceAsync(bool importExchangeModule, CancellationToken cancellationToken)
    {
        var callerCancellation = cancellationToken;
        using var budget = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        budget.CancelAfter(TimeSpan.FromSeconds(30));
        cancellationToken = budget.Token;
        cancellationToken.ThrowIfCancellationRequested();
        DisposeRunspace();

        var runspace = CreateRunspace();
        _engine.Runspace = runspace;

        if (importExchangeModule)
        {
            try
            {
                _engine.ModuleAvailable = await ImportModuleAsync(runspace, PowerShellEngine.ExchangeOnlineModuleName, cancellationToken).ConfigureAwait(false);
            }
            catch (OperationCanceledException ex) when (!callerCancellation.IsCancellationRequested)
            {
                throw new TimeoutException("PowerShell recovery timed out.", ex);
            }
        }

        return true;
    }

    public void DisposeRunspace()
    {
        if (_engine.Runspace == null)
        {
            return;
        }

        try
        {
            _engine.Runspace.Close();
            _engine.Runspace.Dispose();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[PowerShellEngine] Error disposing runspace: {ex.Message}");
        }
        finally
        {
            _engine.Runspace = null;
        }
    }

    private static Runspace CreateRunspace()
    {
        var iss = InitialSessionState.CreateDefault();
        iss.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.RemoteSigned;

        var runspace = RunspaceFactory.CreateRunspace(iss);
        runspace.Open();
        return runspace;
    }

    private static async Task<bool> ImportModuleAsync(Runspace runspace, string moduleName, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(moduleName))
        {
            return false;
        }

        using var ps = System.Management.Automation.PowerShell.Create();
        ps.Runspace = runspace;

        var policy = PowerShellModuleBootstrapPolicy.Resolve(moduleName)
            ?? throw new InvalidOperationException($"Module '{moduleName}' has no approved policy.");
        ps.AddScript(policy.BuildImportScript());

        try
        {
            await InvokeAsync(ps, cancellationToken).ConfigureAwait(false);
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

    private static async Task<bool> EnsurePackageManagementAvailableAsync(Runspace runspace, CancellationToken cancellationToken)
    {
        using var ps = System.Management.Automation.PowerShell.Create();
        ps.Runspace = runspace;

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
            var result = await InvokeAsync(ps, cancellationToken).ConfigureAwait(false);
            return result.FirstOrDefault()?.BaseObject as bool? ?? false;
        }
        catch (RuntimeException ex)
        {
            Debug.WriteLine($"[PowerShellEngine] Failed to prepare PackageManagement: {ex.Message}");
            return false;
        }
    }

    internal static async Task<System.Collections.ObjectModel.Collection<PSObject>> InvokeAsync(
        System.Management.Automation.PowerShell ps, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        using var registration = cancellationToken.Register(() =>
        {
            try { ps.Stop(); }
            catch (Exception ex) { Debug.WriteLine($"[PowerShellEngine] Import cancellation: {ex.Message}"); }
        });
        try
        {
            var result = await Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return ps.Invoke();
            }, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return result;
        }
        catch (PipelineStoppedException) when (cancellationToken.IsCancellationRequested)
        {
            throw new OperationCanceledException(cancellationToken);
        }
    }
}

