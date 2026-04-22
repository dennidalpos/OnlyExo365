using System.Text.Json;
using System.IO.Compression;
using OnlyExo365.Contracts.Diagnostics;
using OnlyExo365.Contracts.Messages;

namespace OnlyExo365.Tests;

public sealed class PersistentLogWriterTests : IDisposable
{
    private readonly string _tempDirectory = Path.Combine(Path.GetTempPath(), "OnlyExo365.Tests", Guid.NewGuid().ToString("N"));

    public PersistentLogWriterTests()
    {
        Directory.CreateDirectory(_tempDirectory);
    }

    [Fact]
    public void Write_PersistsJsonLineWithCorrelationId()
    {
        var writer = new PersistentLogWriter(
            "ui",
            _tempDirectory,
            utcNow: () => new DateTime(2026, 3, 10, 21, 0, 0, DateTimeKind.Utc),
            processId: () => 4242);

        writer.Write(LogLevel.Warning, "Dashboard", "Mailbox sync delayed", "corr-123");

        var filePath = Path.Combine(_tempDirectory, "ui-20260310.log");
        Assert.True(File.Exists(filePath));

        var line = File.ReadAllLines(filePath).Single();
        var payload = JsonSerializer.Deserialize<PersistentLogEntry>(line);

        Assert.NotNull(payload);
        Assert.Equal(LogLevel.Warning, payload.Level);
        Assert.Equal("ui", payload.Component);
        Assert.Equal("Dashboard", payload.Source);
        Assert.Equal("Mailbox sync delayed", payload.Message);
        Assert.Equal("corr-123", payload.CorrelationId);
        Assert.Equal(4242, payload.ProcessId);
    }

    [Fact]
    public void Write_OmitsEmptyCorrelationId()
    {
        var writer = new PersistentLogWriter(
            "supervisor",
            _tempDirectory,
            utcNow: () => new DateTime(2026, 3, 10, 21, 5, 0, DateTimeKind.Utc),
            processId: () => 777);

        writer.Write(LogLevel.Information, "Supervisor", "Worker started", " ");

        var filePath = Path.Combine(_tempDirectory, "supervisor-20260310.log");
        var line = File.ReadAllLines(filePath).Single();
        var payload = JsonSerializer.Deserialize<PersistentLogEntry>(line);

        Assert.NotNull(payload);
        Assert.Null(payload.CorrelationId);
    }

    [Fact]
    public void Write_AppliesRetentionToExpiredLogFiles()
    {
        File.WriteAllText(Path.Combine(_tempDirectory, "ui-20260302.log"), "expired");
        File.WriteAllText(Path.Combine(_tempDirectory, "worker-20260303.log"), "kept");

        var writer = new PersistentLogWriter(
            "ui",
            _tempDirectory,
            utcNow: () => new DateTime(2026, 3, 12, 9, 0, 0, DateTimeKind.Utc),
            retentionDays: 10);

        writer.Write(LogLevel.Information, "Logs", "Retention sweep");

        Assert.False(File.Exists(Path.Combine(_tempDirectory, "ui-20260302.log")));
        Assert.True(File.Exists(Path.Combine(_tempDirectory, "worker-20260303.log")));
        Assert.True(File.Exists(Path.Combine(_tempDirectory, "ui-20260312.log")));
    }

    [Fact]
    public void ExportLogArchive_IncludesPersistedFilesAndManifest()
    {
        File.WriteAllText(Path.Combine(_tempDirectory, "ui-20260312.log"), "{\"message\":\"ui\"}");
        File.WriteAllText(Path.Combine(_tempDirectory, "worker-20260312.log"), "{\"message\":\"worker\"}");

        var archivePath = Path.Combine(_tempDirectory, "exports", "logs.zip");
        var exportedCount = PersistentLogStore.ExportLogArchive(
            archivePath,
            _tempDirectory,
            retentionDays: 21,
            exportedAtUtc: new DateTime(2026, 3, 12, 10, 0, 0, DateTimeKind.Utc));

        Assert.Equal(2, exportedCount);
        Assert.True(File.Exists(archivePath));

        using var archive = ZipFile.OpenRead(archivePath);
        Assert.NotNull(archive.GetEntry("ui-20260312.log"));
        Assert.NotNull(archive.GetEntry("worker-20260312.log"));
        Assert.NotNull(archive.GetEntry(PersistentLogStore.SummaryFileName));

        var manifestEntry = archive.GetEntry("manifest.json");
        Assert.NotNull(manifestEntry);

        using var manifestStream = manifestEntry!.Open();
        using var manifestReader = new StreamReader(manifestStream);
        var manifestJson = manifestReader.ReadToEnd();

        using var manifestDocument = JsonDocument.Parse(manifestJson);
        Assert.Equal(_tempDirectory, manifestDocument.RootElement.GetProperty("logDirectoryPath").GetString());
        Assert.Equal(21, manifestDocument.RootElement.GetProperty("retentionDays").GetInt32());
        Assert.Equal(PersistentLogStore.SummaryFileName, manifestDocument.RootElement.GetProperty("summaryFileName").GetString());
        Assert.Equal(2, manifestDocument.RootElement.GetProperty("files").GetArrayLength());
    }

    [Fact]
    public void Write_GeneratesCollectorFriendlySummaryFile()
    {
        var writer = new PersistentLogWriter(
            "ui",
            _tempDirectory,
            utcNow: () => new DateTime(2026, 3, 12, 10, 15, 0, DateTimeKind.Utc),
            processId: () => 1234);

        writer.Write(LogLevel.Warning, "Dashboard", "Sample warning", "corr-summary");

        var summaryPath = PersistentLogStore.GetSummaryFilePath(_tempDirectory);
        Assert.True(File.Exists(summaryPath));

        var summary = JsonSerializer.Deserialize<PersistentLogSummary>(File.ReadAllText(summaryPath));
        Assert.NotNull(summary);
        Assert.Equal(_tempDirectory, summary.LogDirectoryPath);
        Assert.Equal(1, summary.Files);
        Assert.Equal(1, summary.TotalEntries);
        Assert.Equal(1, summary.Levels.Warning);
        Assert.Single(summary.Components);
        Assert.Single(summary.RecentErrors);
        Assert.Equal("ui", summary.Components[0].Component);
        Assert.Equal("Sample warning", summary.RecentErrors[0].Message);
        Assert.Equal("ui-20260312.log", summary.RecentErrors[0].FileName);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDirectory))
            {
                Directory.Delete(_tempDirectory, recursive: true);
            }
        }
        catch
        {
        }
    }
}

