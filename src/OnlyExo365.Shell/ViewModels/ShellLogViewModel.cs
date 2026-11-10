using System.Collections.ObjectModel;
using System.Windows.Input;
using OnlyExo365.Contracts.Diagnostics;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Shell.Helpers;

namespace OnlyExo365.Shell.ViewModels;

public sealed class ShellLogViewModel : ViewModelBase
{
    private const int MaxLogEntries = 1000;

    private readonly PersistentLogWriter _persistentLogWriter;
    private bool _isVerboseLoggingEnabled;

    public ShellLogViewModel(PersistentLogWriter persistentLogWriter)
    {
        _persistentLogWriter = persistentLogWriter;
        ClearLogsCommand = new RelayCommand(() => LogEntries.Clear());
    }

    public ObservableCollection<LogEntry> LogEntries { get; } = new();

    public bool IsVerboseLoggingEnabled
    {
        get => _isVerboseLoggingEnabled;
        set
        {
            if (SetProperty(ref _isVerboseLoggingEnabled, value))
            {
                AddLog(
                    LogLevel.Information,
                    value
                        ? "Verbose logging enabled - all PowerShell output will be shown"
                        : "Verbose logging disabled - only important messages will be shown");
            }
        }
    }

    public ICommand ClearLogsCommand { get; }

    public void AddLog(LogLevel level, string message, string? source = null, string? correlationId = null)
    {
        if (!IsVerboseLoggingEnabled && level == LogLevel.Verbose)
        {
            return;
        }

        var entry = new LogEntry
        {
            Level = level,
            Message = message,
            Source = source ?? "UI",
            CorrelationId = correlationId
        };

        LogEntries.Add(entry);
        _persistentLogWriter.Write(level, entry.Source ?? "UI", entry.Message, entry.CorrelationId);

        if (LogEntries.Count > MaxLogEntries)
        {
            LogEntries.RemoveAt(0);
        }
    }
}

