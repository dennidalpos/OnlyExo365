using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Dtos;

/// <summary>
/// Richiesta per operazione demo lunga.
/// </summary>
public class DemoOperationRequest
{
    [JsonPropertyName("durationSeconds")]
    public int DurationSeconds { get; set; } = 10;

    [JsonPropertyName("itemCount")]
    public int ItemCount { get; set; } = 10;

    [JsonPropertyName("simulateError")]
    public bool SimulateError { get; set; }

    [JsonPropertyName("errorAtPercent")]
    public int ErrorAtPercent { get; set; } = 50;
}

/// <summary>
/// Risposta operazione demo.
/// </summary>
public class DemoOperationResponse
{
    [JsonPropertyName("processedItems")]
    public int ProcessedItems { get; set; }

    [JsonPropertyName("elapsedSeconds")]
    public double ElapsedSeconds { get; set; }

    [JsonPropertyName("results")]
    public List<DemoItemResult> Results { get; set; } = new();
}

/// <summary>
/// Risultato singolo item demo.
/// </summary>
public class DemoItemResult
{
    [JsonPropertyName("itemId")]
    public int ItemId { get; set; }

    [JsonPropertyName("status")]
    public string Status { get; set; } = "Processed";

    [JsonPropertyName("timestamp")]
    public DateTime Timestamp { get; set; }
}
