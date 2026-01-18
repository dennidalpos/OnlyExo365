namespace ExchangeAdmin.Infrastructure.Ipc;

/// <summary>
/// Stato della connessione al worker.
/// </summary>
public enum WorkerConnectionState
{
    /// <summary>
    /// Worker non avviato.
    /// </summary>
    NotStarted,

    /// <summary>
    /// Worker in fase di avvio.
    /// </summary>
    Starting,

    /// <summary>
    /// Worker avviato, in attesa handshake.
    /// </summary>
    WaitingForHandshake,

    /// <summary>
    /// Worker connesso e pronto.
    /// </summary>
    Connected,

    /// <summary>
    /// Worker in riavvio.
    /// </summary>
    Restarting,

    /// <summary>
    /// Worker fermato normalmente.
    /// </summary>
    Stopped,

    /// <summary>
    /// Worker crashato.
    /// </summary>
    Crashed,

    /// <summary>
    /// Worker non risponde (heartbeat timeout).
    /// </summary>
    Unresponsive
}
