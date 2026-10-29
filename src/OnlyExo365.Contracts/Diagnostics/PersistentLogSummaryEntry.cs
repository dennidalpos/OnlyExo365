using System.Text.Json.Serialization;
using OnlyExo365.Contracts.Messages;

namespace OnlyExo365.Contracts.Diagnostics;

public sealed class PersistentLogSummaryEntry
{
    [JsonPropertyName("timestampUtc")]
    public DateTime TimestampUtc { get; init; }

    [JsonPropertyName("level")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public LogLevel Level { get; init; }

    [JsonPropertyName("component")]
    public string Component { get; init; } = string.Empty;

    [JsonPropertyName("source")]
    public string Source { get; init; } = string.Empty;

    [JsonPropertyName("message")]
    public string Message { get; init; } = string.Empty;

    [JsonPropertyName("correlationId")]
    public string? CorrelationId { get; init; }

    [JsonPropertyName("processId")]
    public int ProcessId { get; init; }

    [JsonPropertyName("fileName")]
    public string FileName { get; init; } = string.Empty;
}

