using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Messages;

             
                                  
              
public enum ErrorCode
{
                                  
    AuthenticationFailed = 100,
    ConditionalAccessBlocked = 101,
    MfaRequired = 102,
    TokenExpired = 103,

                              
    PermissionDenied = 200,
    InsufficientPrivileges = 201,

                                    
    CmdletNotAvailable = 300,
    ModuleNotLoaded = 301,
    InvalidParameter = 302,
    OperationNotSupported = 303,

                             
    Throttling = 400,
    ServiceUnavailable = 401,
    NetworkError = 402,
    Timeout = 403,

                            
    ResourceNotFound = 500,
    ResourceAlreadyExists = 501,

                          
    WorkerNotRunning = 600,
    WorkerCrashed = 601,
    IpcError = 602,

                    
    Unknown = 999
}

             
                                
              
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
