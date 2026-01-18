namespace ExchangeAdmin.Contracts;

/// <summary>
/// Costanti condivise per la comunicazione IPC tra UI e Worker.
/// Questi valori sono usati sia dal client che dal server per garantire coerenza.
/// </summary>
public static class IpcConstants
{
    #region Pipe Names

    /// <summary>
    /// Nome della named pipe per request/response e heartbeat.
    /// </summary>
    public const string PipeName = "ExchangeAdmin_IPC_Main";

    /// <summary>
    /// Nome della named pipe per eventi streaming (one-way dal worker).
    /// </summary>
    public const string EventPipeName = "ExchangeAdmin_IPC_Events";

    #endregion

    #region Timeouts

    /// <summary>
    /// Timeout per connessione alla pipe (ms). Usato durante l'handshake iniziale.
    /// </summary>
    public const int ConnectionTimeoutMs = 10000;

    /// <summary>
    /// Timeout per attesa handshake response dopo connessione pipe (ms).
    /// </summary>
    public const int HandshakeTimeoutMs = 5000;

    /// <summary>
    /// Timeout default per una singola request (ms). Operazioni lunghe possono specificare timeout custom.
    /// </summary>
    public const int RequestTimeoutMs = 300000; // 5 minuti

    /// <summary>
    /// Intervallo tra heartbeat ping (ms).
    /// </summary>
    public const int HeartbeatIntervalMs = 5000;

    /// <summary>
    /// Timeout heartbeat - dopo questo tempo senza pong, il worker è considerato unresponsive (ms).
    /// </summary>
    public const int HeartbeatTimeoutMs = 15000;

    /// <summary>
    /// Grace period prima di dichiarare il worker morto dopo heartbeat timeout (ms).
    /// Permette retry aggiuntivi prima del kill.
    /// </summary>
    public const int HeartbeatGracePeriodMs = 5000;

    /// <summary>
    /// Numero di heartbeat mancati prima di considerare il worker morto.
    /// </summary>
    public const int HeartbeatMissedThreshold = 3;

    #endregion

    #region Buffer & Limits

    /// <summary>
    /// Dimensione buffer pipe in bytes.
    /// </summary>
    public const int PipeBufferSize = 65536;

    /// <summary>
    /// Dimensione massima di un singolo messaggio JSON in bytes (10 MB).
    /// Messaggi più grandi vengono rifiutati per protezione DoS.
    /// </summary>
    public const int MaxMessageSizeBytes = 10 * 1024 * 1024;

    /// <summary>
    /// Numero massimo di eventi streaming per singola request.
    /// Oltre questo limite, nuovi eventi vengono scartati.
    /// </summary>
    public const int MaxEventsPerRequest = 10000;

    /// <summary>
    /// Dimensione massima del buffer di lettura per messaggi parziali (256 KB).
    /// </summary>
    public const int MaxReadBufferSize = 256 * 1024;

    #endregion

    #region Protocol

    /// <summary>
    /// Delimitatore per messaggi JSON su pipe (newline).
    /// Ogni messaggio JSON è terminato da questo carattere.
    /// </summary>
    public const char MessageDelimiter = '\n';

    /// <summary>
    /// Versione minima del protocollo supportata.
    /// </summary>
    public const int ProtocolVersionMajor = 1;

    #endregion

    #region Validation

    /// <summary>
    /// Verifica se una dimensione messaggio è valida.
    /// </summary>
    /// <param name="sizeBytes">Dimensione in bytes.</param>
    /// <returns>True se la dimensione è accettabile.</returns>
    public static bool IsValidMessageSize(long sizeBytes)
        => sizeBytes > 0 && sizeBytes <= MaxMessageSizeBytes;

    /// <summary>
    /// Verifica se il numero di eventi è entro i limiti.
    /// </summary>
    /// <param name="eventCount">Numero di eventi.</param>
    /// <returns>True se entro il limite.</returns>
    public static bool IsEventCountWithinLimit(int eventCount)
        => eventCount >= 0 && eventCount < MaxEventsPerRequest;

    #endregion
}
