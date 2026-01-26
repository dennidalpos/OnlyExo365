using ExchangeAdmin.Worker.Ipc;
using ExchangeAdmin.Worker.PowerShell;

namespace ExchangeAdmin.Worker;

internal class Program
{
    
    
    
    private static async Task EnsureExecutionPolicyAsync()
    {
        try
        {
            Console.WriteLine("[Worker] Checking PowerShell Execution Policy...");

            
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
                Console.WriteLine("[Worker] Warning: Could not check Execution Policy");
                return;
            }

            var currentPolicy = (await process.StandardOutput.ReadToEndAsync()).Trim();
            await process.WaitForExitAsync();

            Console.WriteLine($"[Worker] Current Execution Policy (CurrentUser): {currentPolicy}");

            
            if (currentPolicy != "Bypass" && currentPolicy != "Unrestricted" && currentPolicy != "RemoteSigned")
            {
                Console.WriteLine("[Worker] Setting Execution Policy to RemoteSigned for CurrentUser...");

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
                        Console.WriteLine("[Worker] Execution Policy updated successfully");
                    }
                    else
                    {
                        var error = await setProc.StandardError.ReadToEndAsync();
                        Console.WriteLine($"[Worker] Warning: Could not set Execution Policy: {error}");
                    }
                }
            }
            else
            {
                Console.WriteLine("[Worker] Execution Policy is already acceptable");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Worker] Warning: Failed to check/set Execution Policy: {ex.Message}");
            
        }
    }

    private static async Task<int> Main(string[] args)
    {
        Console.WriteLine($"[Worker] Starting ExchangeAdmin.Worker v1.0.0");
        Console.WriteLine($"[Worker] Process ID: {Environment.ProcessId}");

        try
        {
            
            await EnsureExecutionPolicyAsync();

            
            var psEngine = new PowerShellEngine();
            var initResult = await psEngine.InitializeAsync();

            if (!initResult.Success)
            {
                Console.Error.WriteLine($"[Worker] Failed to initialize PowerShell: {initResult.ErrorMessage}");
                
            }
            else
            {
                Console.WriteLine($"[Worker] PowerShell initialized. Version: {initResult.PowerShellVersion}");
                Console.WriteLine($"[Worker] Module available: {initResult.IsModuleAvailable}");
            }

            
            using var server = new IpcServer(psEngine);
            await server.StartAsync();

            Console.WriteLine("[Worker] IPC server started. Waiting for connections...");

            
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

            Console.WriteLine("[Worker] Shutting down...");
            await server.StopAsync();

            psEngine.Dispose();

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[Worker] Fatal error: {ex}");
            return 1;
        }
    }
}
