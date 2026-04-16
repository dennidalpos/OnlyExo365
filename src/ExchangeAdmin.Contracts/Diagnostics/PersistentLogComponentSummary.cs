using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Diagnostics;

public sealed class PersistentLogComponentSummary
{
    [JsonPropertyName("component")]
    public string Component { get; init; } = string.Empty;

    [JsonPropertyName("entries")]
    public int Entries { get; init; }

    [JsonPropertyName("latestTimestampUtc")]
    public DateTime? LatestTimestampUtc { get; init; }

    [JsonPropertyName("levels")]
    public PersistentLogLevelCounts Levels { get; init; } = new();
}
