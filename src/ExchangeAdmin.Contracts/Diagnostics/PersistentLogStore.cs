using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Contracts.Diagnostics;

public static class PersistentLogStore
{
    public const string SummaryFileName = "observability-summary.json";

    private static readonly JsonSerializerOptions ManifestJsonOptions = new()
    {
        WriteIndented = true
    };

    public static IReadOnlyList<string> GetLogFiles(string? logDirectoryPath = null)
    {
        var directoryPath = ResolveLogDirectoryPath(logDirectoryPath);
        if (!Directory.Exists(directoryPath))
        {
            return [];
        }

        return Directory
            .EnumerateFiles(directoryPath, "*.log", SearchOption.TopDirectoryOnly)
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    public static string GetSummaryFilePath(string? logDirectoryPath = null)
    {
        return Path.Combine(ResolveLogDirectoryPath(logDirectoryPath), SummaryFileName);
    }

    public static int ApplyRetention(string? logDirectoryPath, DateTime referenceUtc, int retentionDays)
    {
        if (retentionDays <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(retentionDays), retentionDays, "Retention days must be greater than zero.");
        }

        var directoryPath = ResolveLogDirectoryPath(logDirectoryPath);
        if (!Directory.Exists(directoryPath))
        {
            return 0;
        }

        var cutoffDateUtc = referenceUtc.Date.AddDays(-(retentionDays - 1));
        var deletedCount = 0;

        foreach (var filePath in Directory.EnumerateFiles(directoryPath, "*.log", SearchOption.TopDirectoryOnly))
        {
            if (!TryGetLogDateUtc(filePath, out var logDateUtc) || logDateUtc >= cutoffDateUtc)
            {
                continue;
            }

            File.Delete(filePath);
            deletedCount++;
        }

        return deletedCount;
    }

    public static int ExportLogArchive(
        string destinationArchivePath,
        string? logDirectoryPath = null,
        int? retentionDays = null,
        DateTime? exportedAtUtc = null)
    {
        if (string.IsNullOrWhiteSpace(destinationArchivePath))
        {
            throw new ArgumentException("Destination archive path is required.", nameof(destinationArchivePath));
        }

        var archivePath = destinationArchivePath.Trim();
        var archiveDirectory = Path.GetDirectoryName(archivePath);
        if (!string.IsNullOrWhiteSpace(archiveDirectory))
        {
            Directory.CreateDirectory(archiveDirectory);
        }

        if (File.Exists(archivePath))
        {
            File.Delete(archivePath);
        }

        var sourceDirectory = ResolveLogDirectoryPath(logDirectoryPath);
        var resolvedRetentionDays = PersistentLogWriter.ResolveRetentionDays(retentionDays);
        var exportTimestampUtc = exportedAtUtc ?? DateTime.UtcNow;
        ApplyRetention(sourceDirectory, exportTimestampUtc, resolvedRetentionDays);
        var files = GetLogFiles(sourceDirectory);
        var summary = BuildSummary(sourceDirectory, exportTimestampUtc);

        using var archive = ZipFile.Open(archivePath, ZipArchiveMode.Create);

        foreach (var filePath in files)
        {
            archive.CreateEntryFromFile(filePath, Path.GetFileName(filePath), CompressionLevel.Optimal);
        }

        var summaryEntry = archive.CreateEntry(SummaryFileName, CompressionLevel.Optimal);
        using (var summaryStream = summaryEntry.Open())
        using (var summaryWriter = new StreamWriter(summaryStream, new UTF8Encoding(false)))
        {
            summaryWriter.Write(JsonSerializer.Serialize(summary, ManifestJsonOptions));
        }

        var manifestEntry = archive.CreateEntry("manifest.json", CompressionLevel.Optimal);
        using var manifestStream = manifestEntry.Open();
        using var manifestWriter = new StreamWriter(manifestStream, new UTF8Encoding(false));
        manifestWriter.Write(JsonSerializer.Serialize(
            new
            {
                exportedAtUtc = exportTimestampUtc,
                logDirectoryPath = sourceDirectory,
                retentionDays = resolvedRetentionDays,
                summaryFileName = SummaryFileName,
                summary,
                files = files.Select(path => new
                {
                    name = Path.GetFileName(path),
                    sizeBytes = new FileInfo(path).Length
                }).ToArray()
            },
            ManifestJsonOptions));

        return files.Count;
    }

    public static PersistentLogSummary BuildSummary(string? logDirectoryPath = null, DateTime? generatedAtUtc = null, int recentErrorLimit = 20)
    {
        if (recentErrorLimit <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(recentErrorLimit), recentErrorLimit, "Recent error limit must be greater than zero.");
        }

        var directoryPath = ResolveLogDirectoryPath(logDirectoryPath);
        var files = GetLogFiles(directoryPath);
        var generatedAt = generatedAtUtc ?? DateTime.UtcNow;
        var levelCounts = new PersistentLogLevelCounts();
        var componentAggregates = new Dictionary<string, ComponentAggregate>(StringComparer.OrdinalIgnoreCase);
        var recentErrors = new List<PersistentLogSummaryEntry>();
        var totalEntries = 0;
        var parseErrors = 0;
        DateTime? earliestTimestampUtc = null;
        DateTime? latestTimestampUtc = null;

        foreach (var filePath in files)
        {
            foreach (var line in File.ReadLines(filePath))
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                PersistentLogEntry? entry;
                try
                {
                    entry = JsonSerializer.Deserialize<PersistentLogEntry>(line);
                }
                catch (JsonException)
                {
                    parseErrors++;
                    continue;
                }

                if (entry is null)
                {
                    parseErrors++;
                    continue;
                }

                totalEntries++;
                levelCounts.Increment(entry.Level);
                earliestTimestampUtc = MinTimestamp(earliestTimestampUtc, entry.TimestampUtc);
                latestTimestampUtc = MaxTimestamp(latestTimestampUtc, entry.TimestampUtc);

                var componentKey = string.IsNullOrWhiteSpace(entry.Component) ? "unknown" : entry.Component;
                if (!componentAggregates.TryGetValue(componentKey, out var aggregate))
                {
                    aggregate = new ComponentAggregate(componentKey);
                    componentAggregates.Add(componentKey, aggregate);
                }

                aggregate.Entries++;
                aggregate.Levels.Increment(entry.Level);
                aggregate.LatestTimestampUtc = MaxTimestamp(aggregate.LatestTimestampUtc, entry.TimestampUtc);

                if (entry.Level < LogLevel.Warning)
                {
                    continue;
                }

                recentErrors.Add(new PersistentLogSummaryEntry
                {
                    TimestampUtc = entry.TimestampUtc,
                    Level = entry.Level,
                    Component = componentKey,
                    Source = entry.Source,
                    Message = entry.Message,
                    CorrelationId = entry.CorrelationId,
                    ProcessId = entry.ProcessId,
                    FileName = Path.GetFileName(filePath)
                });
            }
        }

        return new PersistentLogSummary
        {
            GeneratedAtUtc = generatedAt,
            LogDirectoryPath = directoryPath,
            Files = files.Count,
            TotalEntries = totalEntries,
            ParseErrors = parseErrors,
            EarliestTimestampUtc = earliestTimestampUtc,
            LatestTimestampUtc = latestTimestampUtc,
            Levels = levelCounts,
            Components = componentAggregates.Values
                .OrderBy(component => component.Component, StringComparer.OrdinalIgnoreCase)
                .Select(component => new PersistentLogComponentSummary
                {
                    Component = component.Component,
                    Entries = component.Entries,
                    LatestTimestampUtc = component.LatestTimestampUtc,
                    Levels = component.Levels
                })
                .ToArray(),
            RecentErrors = recentErrors
                .OrderByDescending(entry => entry.TimestampUtc)
                .ThenBy(entry => entry.Component, StringComparer.OrdinalIgnoreCase)
                .Take(recentErrorLimit)
                .ToArray()
        };
    }

    public static string WriteSummaryFile(string? logDirectoryPath = null, DateTime? generatedAtUtc = null, int recentErrorLimit = 20)
    {
        var directoryPath = ResolveLogDirectoryPath(logDirectoryPath);
        Directory.CreateDirectory(directoryPath);

        var summaryPath = GetSummaryFilePath(directoryPath);
        var tempPath = summaryPath + ".tmp";
        var summary = BuildSummary(directoryPath, generatedAtUtc, recentErrorLimit);
        var payload = JsonSerializer.Serialize(summary, ManifestJsonOptions);

        File.WriteAllText(tempPath, payload, new UTF8Encoding(false));
        File.Move(tempPath, summaryPath, overwrite: true);
        return summaryPath;
    }

    private static string ResolveLogDirectoryPath(string? logDirectoryPath)
    {
        return string.IsNullOrWhiteSpace(logDirectoryPath)
            ? PersistentLogWriter.GetDefaultLogDirectoryPath()
            : logDirectoryPath;
    }

    private static bool TryGetLogDateUtc(string filePath, out DateTime logDateUtc)
    {
        var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
        var separatorIndex = fileNameWithoutExtension.LastIndexOf('-');
        if (separatorIndex < 0 || separatorIndex == fileNameWithoutExtension.Length - 1)
        {
            logDateUtc = default;
            return false;
        }

        var dateToken = fileNameWithoutExtension[(separatorIndex + 1)..];
        return DateTime.TryParseExact(
            dateToken,
            "yyyyMMdd",
            CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
            out logDateUtc);
    }

    private static DateTime? MinTimestamp(DateTime? current, DateTime candidate)
        => !current.HasValue || candidate < current.Value ? candidate : current;

    private static DateTime? MaxTimestamp(DateTime? current, DateTime candidate)
        => !current.HasValue || candidate > current.Value ? candidate : current;

    private sealed class ComponentAggregate
    {
        public ComponentAggregate(string component)
        {
            Component = component;
        }

        public string Component { get; }

        public int Entries { get; set; }

        public DateTime? LatestTimestampUtc { get; set; }

        public PersistentLogLevelCounts Levels { get; } = new();
    }
}
