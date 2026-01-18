using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Messages;

/// <summary>
/// Codici di errore normalizzati.
/// </summary>
public enum ErrorCode
{
    // Authentication errors (1xx)
    AuthenticationFailed = 100,
    ConditionalAccessBlocked = 101,
    MfaRequired = 102,
    TokenExpired = 103,

    // Permission errors (2xx)
    PermissionDenied = 200,
    InsufficientPrivileges = 201,

    // Cmdlet/Operation errors (3xx)
    CmdletNotAvailable = 300,
    ModuleNotLoaded = 301,
    InvalidParameter = 302,
    OperationNotSupported = 303,

    // Transient errors (4xx)
    Throttling = 400,
    ServiceUnavailable = 401,
    NetworkError = 402,
    Timeout = 403,

    // Resource errors (5xx)
    ResourceNotFound = 500,
    ResourceAlreadyExists = 501,

    // Worker errors (6xx)
    WorkerNotRunning = 600,
    WorkerCrashed = 601,
    IpcError = 602,

    // Unknown (9xx)
    Unknown = 999
}

/// <summary>
/// Errore normalizzato per IPC.
/// </summary>
public class NormalizedErrorDto
{
    [JsonPropertyName("code")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public ErrorCode Code { get; set; }

    [JsonPropertyName("message")]
    public string Message { get; set; } = string.Empty;

    [JsonPropertyName("details")]
    public string? Details { get; set; }

    [JsonPropertyName("isTransient")]
    public bool IsTransient { get; set; }

    [JsonPropertyName("retryAfterSeconds")]
    public int? RetryAfterSeconds { get; set; }

    [JsonPropertyName("innerException")]
    public string? InnerException { get; set; }

    [JsonPropertyName("stackTrace")]
    public string? StackTrace { get; set; }
}
