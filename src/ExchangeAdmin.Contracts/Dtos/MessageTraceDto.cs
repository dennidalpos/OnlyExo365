using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Dtos;

public class GetMessageTraceRequest
{
    [JsonPropertyName("senderAddress")]
    public string? SenderAddress { get; set; }

    [JsonPropertyName("recipientAddress")]
    public string? RecipientAddress { get; set; }

    [JsonPropertyName("startDate")]
    public DateTime StartDate { get; set; }

    [JsonPropertyName("endDate")]
    public DateTime EndDate { get; set; }

    [JsonPropertyName("pageSize")]
    public int PageSize { get; set; } = 100;

    [JsonPropertyName("page")]
    public int Page { get; set; } = 1;
}

public class GetMessageTraceResponse
{
    [JsonPropertyName("messages")]
    public List<MessageTraceItemDto> Messages { get; set; } = new();

    [JsonPropertyName("totalCount")]
    public int TotalCount { get; set; }

    [JsonPropertyName("hasMore")]
    public bool HasMore { get; set; }

    [JsonPropertyName("page")]
    public int Page { get; set; }

    [JsonPropertyName("pageSize")]
    public int PageSize { get; set; }
}

public class GetMessageTraceDetailsRequest
{
    [JsonPropertyName("messageTraceId")]
    public string MessageTraceId { get; set; } = string.Empty;

    [JsonPropertyName("recipientAddress")]
    public string RecipientAddress { get; set; } = string.Empty;
}

public class GetMessageTraceDetailsResponse
{
    [JsonPropertyName("messageTraceId")]
    public string MessageTraceId { get; set; } = string.Empty;

    [JsonPropertyName("events")]
    public List<MessageTraceDetailEventDto> Events { get; set; } = new();
}

public class MessageTraceDetailEventDto
{
    [JsonPropertyName("date")]
    public DateTime? Date { get; set; }

    [JsonPropertyName("event")]
    public string Event { get; set; } = string.Empty;

    [JsonPropertyName("action")]
    public string Action { get; set; } = string.Empty;

    [JsonPropertyName("detail")]
    public string Detail { get; set; } = string.Empty;

    [JsonPropertyName("data")]
    public string Data { get; set; } = string.Empty;
}

public class MessageTraceItemDto
{
    [JsonPropertyName("messageId")]
    public string MessageId { get; set; } = string.Empty;

    [JsonPropertyName("messageTraceId")]
    public string MessageTraceId { get; set; } = string.Empty;

    [JsonPropertyName("senderAddress")]
    public string SenderAddress { get; set; } = string.Empty;

    [JsonPropertyName("recipientAddress")]
    public string RecipientAddress { get; set; } = string.Empty;

    [JsonPropertyName("subject")]
    public string Subject { get; set; } = string.Empty;

    [JsonPropertyName("status")]
    public string Status { get; set; } = string.Empty;

    [JsonPropertyName("received")]
    public DateTime? Received { get; set; }

    [JsonPropertyName("size")]
    public long? Size { get; set; }
}
