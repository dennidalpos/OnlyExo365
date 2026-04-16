using System.Text;
using System.Windows;
using System.Windows.Threading;
using ExchangeAdmin.Contracts.Diagnostics;
using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Presentation.Services;

internal sealed class GlobalExceptionMonitor : IDisposable
{
    private readonly PersistentLogWriter _persistentLogWriter;
    private readonly Action<string, string, string?> _showError;
    private readonly Action<int> _shutdown;
    private bool _isRegistered;
    private bool _fatalShutdownRequested;

    public GlobalExceptionMonitor(
        PersistentLogWriter persistentLogWriter,
        Action<string, string, string?>? showError = null,
        Action<int>? shutdown = null)
    {
        _persistentLogWriter = persistentLogWriter ?? throw new ArgumentNullException(nameof(persistentLogWriter));
        _showError = showError ?? ErrorDialogService.ShowError;
        _shutdown = shutdown ?? (exitCode => System.Windows.Application.Current?.Shutdown(exitCode));
    }

    public void Register(System.Windows.Application application)
    {
        ArgumentNullException.ThrowIfNull(application);

        if (_isRegistered)
        {
            return;
        }

        application.DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
        _isRegistered = true;
    }

    public void HandleDispatcherUnhandledException(Exception exception)
    {
        ArgumentNullException.ThrowIfNull(exception);

        LogFatal("DispatcherUnhandledException", exception, isTerminating: true);

        if (_fatalShutdownRequested)
        {
            return;
        }

        _fatalShutdownRequested = true;
        _showError(
            "Unexpected Application Error",
            "A fatal UI error occurred and the application will be closed.",
            exception.Message);
        _shutdown(-1);
    }

    public void HandleCurrentDomainUnhandledException(Exception exception, bool isTerminating)
    {
        ArgumentNullException.ThrowIfNull(exception);
        LogFatal("AppDomainUnhandledException", exception, isTerminating);
    }

    public void HandleUnobservedTaskException(AggregateException exception)
    {
        ArgumentNullException.ThrowIfNull(exception);
        LogFatal("TaskSchedulerUnobservedTaskException", exception, isTerminating: false);
    }

    private void OnDispatcherUnhandledException(object? sender, DispatcherUnhandledExceptionEventArgs e)
    {
        HandleDispatcherUnhandledException(e.Exception);
        e.Handled = true;
    }

    private void OnUnhandledException(object? sender, UnhandledExceptionEventArgs e)
    {
        var exception = e.ExceptionObject as Exception
            ?? new InvalidOperationException($"Unhandled non-exception object: {e.ExceptionObject}");

        HandleCurrentDomainUnhandledException(exception, e.IsTerminating);
    }

    private void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        HandleUnobservedTaskException(e.Exception);
        e.SetObserved();
    }

    private void LogFatal(string source, Exception exception, bool isTerminating)
    {
        var details = BuildExceptionDetails(exception, isTerminating);
        _persistentLogWriter.Write(LogLevel.Error, source, details);
    }

    private static string BuildExceptionDetails(Exception exception, bool isTerminating)
    {
        var builder = new StringBuilder();
        builder.Append("Unhandled exception captured");
        builder.Append(" | terminating=");
        builder.Append(isTerminating ? "true" : "false");
        builder.AppendLine();
        builder.Append(exception);
        return builder.ToString();
    }

    public void Dispose()
    {
        if (!_isRegistered)
        {
            return;
        }

        if (System.Windows.Application.Current != null)
        {
            System.Windows.Application.Current.DispatcherUnhandledException -= OnDispatcherUnhandledException;
        }

        AppDomain.CurrentDomain.UnhandledException -= OnUnhandledException;
        TaskScheduler.UnobservedTaskException -= OnUnobservedTaskException;
        _isRegistered = false;
    }
}
