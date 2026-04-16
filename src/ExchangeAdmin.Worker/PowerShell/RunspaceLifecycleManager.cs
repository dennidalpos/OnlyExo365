using System.Diagnostics;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace ExchangeAdmin.Worker.PowerShell;

internal sealed class RunspaceLifecycleManager
{
    private readonly PowerShellEngine _engine;

    public RunspaceLifecycleManager(PowerShellEngine engine)
    {
        _engine = engine;
    }

    public async Task<PowerShellInitResult> InitializeAsync()
    {
        try
        {
            Debug.WriteLine("[PowerShellEngine] Initializing...");

            var runspace = CreateRunspace();
            _engine.Runspace = runspace;

            Debug.WriteLine($"[PowerShellEngine] Runspace opened, state: {runspace.RunspaceStateInfo.State}");

            using var ps = System.Management.Automation.PowerShell.Create();
            ps.Runspace = runspace;
            ps.AddScript("$PSVersionTable.PSVersion.ToString()");

            var versionResult = await Task.Run(() => ps.Invoke()).ConfigureAwait(false);
            _engine.PowerShellVersionValue = versionResult.FirstOrDefault()?.ToString() ?? "Unknown";

            Debug.WriteLine($"[PowerShellEngine] PowerShell version: {_engine.PowerShellVersionValue}");

            ps.Commands.Clear();
            ps.AddScript($"Get-Module -ListAvailable -Name {PowerShellEngine.ExchangeOnlineModuleName} | Select-Object -First 1");

            var moduleResult = await Task.Run(() => ps.Invoke()).ConfigureAwait(false);
            _engine.ModuleAvailable = moduleResult.Any();

            Debug.WriteLine($"[PowerShellEngine] ExchangeOnlineManagement module available: {_engine.ModuleAvailable}");

            if (_engine.ModuleAvailable)
            {
                var packageManagementReady = await EnsurePackageManagementAvailableAsync(runspace).ConfigureAwait(false);
                if (!packageManagementReady)
                {
                    Debug.WriteLine("[PowerShellEngine] PackageManagement module was not explicitly loaded. Continuing with best-effort import.");
                }

                var imported = await ImportModuleAsync(runspace, PowerShellEngine.ExchangeOnlineModuleName, stopOnError: true).ConfigureAwait(false);
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

    public async Task<bool> RecreateRunspaceAsync(bool importExchangeModule)
    {
        DisposeRunspace();

        var runspace = CreateRunspace();
        _engine.Runspace = runspace;

        if (importExchangeModule)
        {
            _ = await ImportModuleAsync(runspace, PowerShellEngine.ExchangeOnlineModuleName, stopOnError: false).ConfigureAwait(false);
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

    private static async Task<bool> ImportModuleAsync(Runspace runspace, string moduleName, bool stopOnError)
    {
        if (string.IsNullOrWhiteSpace(moduleName))
        {
            return false;
        }

        using var ps = System.Management.Automation.PowerShell.Create();
        ps.Runspace = runspace;

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

    private static async Task<bool> EnsurePackageManagementAvailableAsync(Runspace runspace)
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
            var result = await Task.Run(() => ps.Invoke()).ConfigureAwait(false);
            return result.FirstOrDefault()?.BaseObject as bool? ?? false;
        }
        catch (RuntimeException ex)
        {
            Debug.WriteLine($"[PowerShellEngine] Failed to prepare PackageManagement: {ex.Message}");
            return false;
        }
    }
}
