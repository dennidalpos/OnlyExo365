using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Messages;

             
                                                   
              
public static class JsonMessageSerializer
{
    private static readonly JsonSerializerOptions Options = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = false,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters =
        {
            new JsonStringEnumConverter(JsonNamingPolicy.CamelCase)
        }
    };

    public static string Serialize<T>(T message)
    {
        return JsonSerializer.Serialize(message, Options);
    }

    public static T? Deserialize<T>(string json)
    {
        return JsonSerializer.Deserialize<T>(json, Options);
    }

    public static object? Deserialize(string json, Type type)
    {
        return JsonSerializer.Deserialize(json, type, Options);
    }

    public static IpcMessage? DeserializeMessage(string json)
    {
                                                                   
        var baseMessage = JsonSerializer.Deserialize<IpcMessage>(json, Options);
        if (baseMessage == null) return null;

                                               
        return baseMessage.Type switch
        {
            MessageType.HandshakeRequest => JsonSerializer.Deserialize<HandshakeRequest>(json, Options),
            MessageType.HandshakeResponse => JsonSerializer.Deserialize<HandshakeResponse>(json, Options),
            MessageType.Request => JsonSerializer.Deserialize<RequestEnvelope>(json, Options),
            MessageType.Response => JsonSerializer.Deserialize<ResponseEnvelope>(json, Options),
            MessageType.Event => JsonSerializer.Deserialize<EventEnvelope>(json, Options),
            MessageType.CancelRequest => JsonSerializer.Deserialize<CancelRequest>(json, Options),
            MessageType.HeartbeatPing => JsonSerializer.Deserialize<HeartbeatPing>(json, Options),
            MessageType.HeartbeatPong => JsonSerializer.Deserialize<HeartbeatPong>(json, Options),
            _ => baseMessage
        };
    }

    public static T? ExtractPayload<T>(JsonElement? payload)
    {
        if (payload == null || payload.Value.ValueKind == JsonValueKind.Null)
            return default;

        return JsonSerializer.Deserialize<T>(payload.Value.GetRawText(), Options);
    }

    public static JsonElement ToJsonElement<T>(T value)
    {
        var json = JsonSerializer.Serialize(value, Options);
        using var document = JsonDocument.Parse(json);
        return document.RootElement.Clone();
    }
}
