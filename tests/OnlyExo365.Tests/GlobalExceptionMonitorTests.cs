using System.Text.Json;
using OnlyExo365.Contracts.Diagnostics;
using OnlyExo365.Shell.Services;

namespace OnlyExo365.Tests;

public sealed class GlobalExceptionMonitorTests : IDisposable
{
    private readonly string _tempDirectory;

    public GlobalExceptionMonitorTests()
    {
        _tempDirectory = Path.Combine(Path.GetTempPath(), "OnlyExo365.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDirectory);
    }

    [WpfFact]
    public void HandleDispatcherUnhandledException_LogsFatalError_ShowsDialog_AndRequestsShutdown()
    {
        var writer = new PersistentLogWriter("ui", _tempDirectory);
        var dialogs = new List<(string Title, string Message, string? Details)>();
        var shutdownCodes = new List<int>();
        using var monitor = new GlobalExceptionMonitor(
            writer,
            (title, message, details) => dialogs.Add((title, message, details)),
            exitCode => shutdownCodes.Add(exitCode));
        var exception = new InvalidOperationException("Boom from UI");

        monitor.HandleDispatcherUnhandledException(exception);

        var payload = ReadSinglePayload();
        Assert.Equal("DispatcherUnhandledException", payload.Source);
        Assert.Contains("Unhandled exception captured | terminating=true", payload.Message);
        Assert.Contains("Boom from UI", payload.Message);
        Assert.Single(dialogs);
        Assert.Equal("Unexpected Application Error", dialogs[0].Title);
        Assert.Equal("A fatal UI error occurred and the application will be closed.", dialogs[0].Message);
        Assert.Equal("Boom from UI", dialogs[0].Details);
        Assert.Single(shutdownCodes);
        Assert.Equal(-1, shutdownCodes[0]);
    }

    [WpfFact]
    public void HandleCurrentDomainUnhandledException_LogsFatalErrorWithoutDialog()
    {
        var writer = new PersistentLogWriter("ui", _tempDirectory);
        var dialogs = new List<string>();
        using var monitor = new GlobalExceptionMonitor(
            writer,
            (title, message, details) => dialogs.Add(title),
            _ => throw new InvalidOperationException("Shutdown should not be requested."));
        var exception = new ApplicationException("Background thread crash");

        monitor.HandleCurrentDomainUnhandledException(exception, isTerminating: true);

        var payload = ReadSinglePayload();
        Assert.Equal("AppDomainUnhandledException", payload.Source);
        Assert.Contains("terminating=true", payload.Message);
        Assert.Contains("Background thread crash", payload.Message);
        Assert.Empty(dialogs);
    }

    [WpfFact]
    public void HandleUnobservedTaskException_LogsAndDoesNotRequestShutdown()
    {
        var writer = new PersistentLogWriter("ui", _tempDirectory);
        var shutdownCodes = new List<int>();
        using var monitor = new GlobalExceptionMonitor(
            writer,
            (_, _, _) => { },
            exitCode => shutdownCodes.Add(exitCode));
        var exception = new AggregateException(new InvalidOperationException("Unobserved task failure"));

        monitor.HandleUnobservedTaskException(exception);

        var payload = ReadSinglePayload();
        Assert.Equal("TaskSchedulerUnobservedTaskException", payload.Source);
        Assert.Contains("terminating=false", payload.Message);
        Assert.Contains("Unobserved task failure", payload.Message);
        Assert.Empty(shutdownCodes);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDirectory))
        {
            Directory.Delete(_tempDirectory, recursive: true);
        }
    }

    private PersistentLogEntry ReadSinglePayload()
    {
        var logFile = Directory.GetFiles(_tempDirectory, "*.log", SearchOption.TopDirectoryOnly).Single();
        var json = File.ReadAllText(logFile);
        return JsonSerializer.Deserialize<PersistentLogEntry>(json)!;
    }
}

