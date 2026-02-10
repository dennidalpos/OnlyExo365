using ExchangeAdmin.Worker.Ipc;
using ExchangeAdmin.Worker.PowerShell;

namespace ExchangeAdmin.Worker;

internal class Program
{
                 
                                                                                               
                  
    private static async Task EnsureExecutionPolicyAsync()
    {
        const string source = "Worker";
        try
        {
            ConsoleLogger.Info(source, "Checking PowerShell Execution Policy...");

            var checkProcess = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "pwsh",
                Arguments = "-NoProfile -Command \"Get-ExecutionPolicy -Scope CurrentUser\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };

            using var process = System.Diagnostics.Process.Start(checkProcess);
            if (process == null)
            {
                ConsoleLogger.Warning(source, "Could not check Execution Policy");
                return;
            }

            var currentPolicy = (await process.StandardOutput.ReadToEndAsync()).Trim();
            await process.WaitForExitAsync();

            ConsoleLogger.Debug(source, $"Current Execution Policy (CurrentUser): {currentPolicy}");

            if (currentPolicy != "Bypass" && currentPolicy != "Unrestricted" && currentPolicy != "RemoteSigned")
            {
                ConsoleLogger.Info(source, "Setting Execution Policy to RemoteSigned for CurrentUser...");

                var setProcess = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "pwsh",
                    Arguments = "-NoProfile -Command \"Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using var setProc = System.Diagnostics.Process.Start(setProcess);
                if (setProc != null)
                {
                    await setProc.WaitForExitAsync();
                    if (setProc.ExitCode == 0)
                    {
                        ConsoleLogger.Success(source, "Execution Policy updated successfully");
                    }
                    else
                    {
                        var error = await setProc.StandardError.ReadToEndAsync();
                        ConsoleLogger.Warning(source, $"Could not set Execution Policy: {error}");
                    }
                }
            }
            else
            {
                ConsoleLogger.Success(source, "Execution Policy is already acceptable");
            }
        }
        catch (Exception ex)
        {
            ConsoleLogger.Warning(source, $"Failed to check/set Execution Policy: {ex.Message}");
        }
    }

    private static async Task<int> Main(string[] args)
    {
        const string source = "Worker";
        ConsoleLogger.Info(source, "Starting ExchangeAdmin.Worker v1.0.1");
        ConsoleLogger.Debug(source, $"Process ID: {Environment.ProcessId}");

        try
        {
            await EnsureExecutionPolicyAsync();

            var psEngine = new PowerShellEngine();
            var initResult = await psEngine.InitializeAsync();

            if (!initResult.Success)
            {
                ConsoleLogger.Error(source, $"Failed to initialize PowerShell: {initResult.ErrorMessage}");
            }
            else
            {
                ConsoleLogger.Success(source, $"PowerShell initialized. Version: {initResult.PowerShellVersion}");
                ConsoleLogger.Info(source, $"Module available: {initResult.IsModuleAvailable}");
            }

            using var server = new IpcServer(psEngine);
            await server.StartAsync();

            ConsoleLogger.Success(source, "IPC server started. Waiting for connections...");

            using var shutdownEvent = new ManualResetEventSlim(false);

            Console.CancelKeyPress += (s, e) =>
            {
                e.Cancel = true;
                shutdownEvent.Set();
            };

            AppDomain.CurrentDomain.ProcessExit += (s, e) =>
            {
                shutdownEvent.Set();
            };

            while (!shutdownEvent.Wait(1000))
            {
            }

            ConsoleLogger.Warning(source, "Shutting down...");
            await server.StopAsync();

            psEngine.Dispose();

            return 0;
        }
        catch (Exception ex)
        {
            ConsoleLogger.Error(source, $"Fatal error: {ex}");
            return 1;
        }
    }
}
