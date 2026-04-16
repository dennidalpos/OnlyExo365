using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Diagnostics;

public sealed class PersistentLogSummary
{
    [JsonPropertyName("generatedAtUtc")]
    public DateTime GeneratedAtUtc { get; init; }

    [JsonPropertyName("logDirectoryPath")]
    public string LogDirectoryPath { get; init; } = string.Empty;

    [JsonPropertyName("files")]
    public int Files { get; init; }

    [JsonPropertyName("totalEntries")]
    public int TotalEntries { get; init; }

    [JsonPropertyName("parseErrors")]
    public int ParseErrors { get; init; }

    [JsonPropertyName("earliestTimestampUtc")]
    public DateTime? EarliestTimestampUtc { get; init; }

    [JsonPropertyName("latestTimestampUtc")]
    public DateTime? LatestTimestampUtc { get; init; }

    [JsonPropertyName("levels")]
    public PersistentLogLevelCounts Levels { get; init; } = new();

    [JsonPropertyName("components")]
    public IReadOnlyList<PersistentLogComponentSummary> Components { get; init; } = [];

    [JsonPropertyName("recentErrors")]
    public IReadOnlyList<PersistentLogSummaryEntry> RecentErrors { get; init; } = [];
}
