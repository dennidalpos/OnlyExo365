using System.Management.Automation.Runspaces;
using OnlyExo365.Contracts.Diagnostics;
using OnlyExo365.Worker;
using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Tests;

internal sealed class LocalPowerShellFixture : IDisposable
{
    private readonly string _logDirectory = Path.Combine(Path.GetTempPath(), "OnlyExo365.Tests", Guid.NewGuid().ToString("N"));

    public LocalPowerShellFixture()
    {
        ConsoleLogger.SetPersistentLogWriterForTesting(new PersistentLogWriter("probe", _logDirectory));
        var runspace = RunspaceFactory.CreateRunspace(InitialSessionState.CreateDefault());
        runspace.Open();
        Engine = new PowerShellEngine { Runspace = runspace, Initialized = true };
    }

    public PowerShellEngine Engine { get; }

    public async Task DefineCommandsAsync(string script)
    {
        var result = await Engine.ExecuteAsync(script);
        Assert.True(result.Success, result.ErrorMessage);
    }

    public void Dispose()
    {
        Engine.Dispose();
        ConsoleLogger.ResetPersistentLogWriterForTesting();
        if (Directory.Exists(_logDirectory))
        {
            Directory.Delete(_logDirectory, recursive: true);
        }
    }
}
