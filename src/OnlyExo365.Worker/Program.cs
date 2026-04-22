using OnlyExo365.Contracts;
using OnlyExo365.Worker.Ipc;
using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Worker;

internal class Program
{
    private static async Task LogExecutionPolicyStatusAsync()
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

            if (!string.Equals(currentPolicy, "Bypass", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(currentPolicy, "Unrestricted", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(currentPolicy, "RemoteSigned", StringComparison.OrdinalIgnoreCase))
            {
                ConsoleLogger.Warning(source, "CurrentUser Execution Policy is not in the supported set (RemoteSigned, Unrestricted, Bypass).");
                ConsoleLogger.Warning(source, "Run manually in PowerShell 7 if required: Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned");
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
        ConsoleLogger.Info(source, $"Starting OnlyExo365.Worker {ProductInfo.DisplayVersion}");
        ConsoleLogger.Debug(source, $"Process ID: {Environment.ProcessId}");

        try
        {
            await LogExecutionPolicyStatusAsync();

            var exchangeConfiguration = ExchangeOnlineConfiguration.FromEnvironmentVariables(Environment.GetEnvironmentVariable);
            foreach (var validationError in exchangeConfiguration.Validate())
            {
                ConsoleLogger.Warning(source, $"Exchange configuration: {validationError}");
            }

            var psEngine = new PowerShellEngine(exchangeConfiguration);
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

            var ipcSessionContext = IpcSessionContext.CreateForCurrentProcess();
            var ipcSessionToken = Environment.GetEnvironmentVariable(IpcConstants.SessionTokenEnvironmentVariable);

            if (string.IsNullOrWhiteSpace(ipcSessionToken))
            {
                ConsoleLogger.Error(source, $"Missing required environment variable: {IpcConstants.SessionTokenEnvironmentVariable}");
                psEngine.Dispose();
                return 1;
            }

            using var server = new IpcServer(psEngine, ipcSessionContext, ipcSessionToken);
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

