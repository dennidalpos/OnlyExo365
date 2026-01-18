namespace ExchangeAdmin.Contracts.Messages;

/// <summary>
/// Tipi di evento streaming.
/// </summary>
public enum EventType
{
    /// <summary>
    /// Log message (verbose, info, warning, error).
    /// </summary>
    Log,

    /// <summary>
    /// Progress update (percentuale completamento).
    /// </summary>
    Progress,

    /// <summary>
    /// Partial output (risultati parziali durante operazione lunga).
    /// </summary>
    PartialOutput
}
