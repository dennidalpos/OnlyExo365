namespace ExchangeAdmin.Contracts.Messages;

/// <summary>
/// Tipi di messaggio IPC.
/// </summary>
public enum MessageType
{
    // Handshake
    HandshakeRequest,
    HandshakeResponse,

    // Request/Response
    Request,
    Response,

    // Streaming events
    Event,

    // Control
    CancelRequest,
    HeartbeatPing,
    HeartbeatPong
}
