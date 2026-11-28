using System.Diagnostics;
using System.Text;
using System.Text.Json;
using OnlyExo365.Contracts.Messages;

namespace OnlyExo365.Contracts.Diagnostics;

public sealed class PersistentLogWriter
{
    public const int DefaultRetentionDays = 14;
    public const string RetentionDaysEnvironmentVariableName = "ONLYEXO365_LOG_RETENTION_DAYS";
    private static readonly TimeSpan SummaryRefreshInterval = TimeSpan.FromSeconds(5);

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = false
    };

    private readonly object _sync = new();
    private readonly string _component;
    private readonly string _logDirectoryPath;
    private readonly Func<DateTime> _utcNow;
    private readonly Func<int> _processId;
    private readonly int _retentionDays;
    private DateTime _lastRetentionSweepDateUtc = DateTime.MinValue;
    private DateTime _lastSummaryRefreshUtc = DateTime.MinValue;

    public PersistentLogWriter(
        string component,
        string? logDirectoryPath = null,
        Func<DateTime>? utcNow = null,
        Func<int>? processId = null,
        int? retentionDays = null)
    {
        if (string.IsNullOrWhiteSpace(component))
        {
            throw new ArgumentException("Component is required.", nameof(component));
        }

        _component = SanitizeSegment(component);
        _logDirectoryPath = string.IsNullOrWhiteSpace(logDirectoryPath)
            ? GetDefaultLogDirectoryPath()
            : logDirectoryPath;
        _utcNow = utcNow ?? (() => DateTime.UtcNow);
        _processId = processId ?? (() => Environment.ProcessId);
        _retentionDays = ResolveRetentionDays(retentionDays);
    }

    public string CurrentLogFilePath => GetLogFilePath(_utcNow());
    public string LogDirectoryPath => _logDirectoryPath;
    public int RetentionDays => _retentionDays;

    public void Write(LogLevel level, string source, string message, string? correlationId = null)
    {
        if (string.IsNullOrWhiteSpace(source))
        {
            throw new ArgumentException("Source is required.", nameof(source));
        }

        if (message == null)
        {
            throw new ArgumentNullException(nameof(message));
        }

        var timestamp = _utcNow();
        var entry = new PersistentLogEntry
        {
            TimestampUtc = timestamp,
            Level = level,
            Component = _component,
            Source = source.Trim(),
            Message = message,
            CorrelationId = string.IsNullOrWhiteSpace(correlationId) ? null : correlationId.Trim(),
            ProcessId = _processId()
        };

        WriteEntry(entry, timestamp);
    }

    internal void WriteEntry(PersistentLogEntry entry, DateTime timestampUtc)
    {
        try
        {
            lock (_sync)
            {
                Directory.CreateDirectory(_logDirectoryPath);
                ApplyRetentionIfRequired(timestampUtc);

                var payload = JsonSerializer.Serialize(entry, JsonOptions);
                var filePath = GetLogFilePath(timestampUtc);
                File.AppendAllText(filePath, payload + Environment.NewLine, new UTF8Encoding(false));
                RefreshSummaryIfRequired(timestampUtc);
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            Trace.WriteLine($"[PersistentLogWriter] Unable to persist log entry: {ex.Message}");
        }
    }

    public static string GetDefaultLogDirectoryPath()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OnlyExo365",
            "logs");
    }

    public static int ResolveRetentionDays(int? overrideRetentionDays = null)
    {
        if (overrideRetentionDays.HasValue)
        {
            return overrideRetentionDays.Value > 0
                ? overrideRetentionDays.Value
                : DefaultRetentionDays;
        }

        var fromEnvironment = Environment.GetEnvironmentVariable(RetentionDaysEnvironmentVariableName);
        return int.TryParse(fromEnvironment, out var parsedDays) && parsedDays > 0
            ? parsedDays
            : DefaultRetentionDays;
    }

    private string GetLogFilePath(DateTime timestampUtc)
    {
        var fileName = $"{_component}-{timestampUtc:yyyyMMdd}.log";
        return Path.Combine(_logDirectoryPath, fileName);
    }

    private void ApplyRetentionIfRequired(DateTime timestampUtc)
    {
        var sweepDateUtc = timestampUtc.Date;
        if (_lastRetentionSweepDateUtc == sweepDateUtc)
        {
            return;
        }

        PersistentLogStore.ApplyRetention(_logDirectoryPath, sweepDateUtc, _retentionDays);
        _lastRetentionSweepDateUtc = sweepDateUtc;
    }

    private void RefreshSummaryIfRequired(DateTime timestampUtc)
    {
        if (_lastSummaryRefreshUtc != DateTime.MinValue &&
            timestampUtc - _lastSummaryRefreshUtc < SummaryRefreshInterval)
        {
            return;
        }

        PersistentLogStore.WriteSummaryFile(_logDirectoryPath, timestampUtc);
        _lastSummaryRefreshUtc = timestampUtc;
    }

    private static string SanitizeSegment(string value)
    {
        var invalidChars = Path.GetInvalidFileNameChars();
        var builder = new StringBuilder(value.Length);

        foreach (var ch in value.Trim().ToLowerInvariant())
        {
            builder.Append(invalidChars.Contains(ch) || char.IsWhiteSpace(ch) ? '-' : ch);
        }

        return builder.ToString().Trim('-');
    }
}

