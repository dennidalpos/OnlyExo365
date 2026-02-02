using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Domain.Errors;

             
                                       
              
public class NormalizedError
{
    public ErrorCode Code { get; init; }
    public string Message { get; init; } = string.Empty;
    public string? Details { get; init; }
    public bool IsTransient { get; init; }
    public TimeSpan? RetryAfter { get; init; }
    public Exception? OriginalException { get; init; }

    private NormalizedError() { }

                 
                                                    
                  
    public static NormalizedError FromException(Exception ex)
    {
        var (category, isTransient) = ErrorTaxonomy.Classify(ex.Message, ex.GetType().Name);
        var retryAfter = ErrorTaxonomy.ExtractRetryAfter(ex.Message);

        var code = category switch
        {
            ErrorCategory.Authentication when ex.Message.Contains("MFA", StringComparison.OrdinalIgnoreCase)
                => ErrorCode.MfaRequired,
            ErrorCategory.Authentication when ex.Message.Contains("Conditional", StringComparison.OrdinalIgnoreCase)
                => ErrorCode.ConditionalAccessBlocked,
            ErrorCategory.Authentication when ex.Message.Contains("expired", StringComparison.OrdinalIgnoreCase)
                => ErrorCode.TokenExpired,
            ErrorCategory.Authentication => ErrorCode.AuthenticationFailed,

            ErrorCategory.Permission => ErrorCode.PermissionDenied,

            ErrorCategory.Operation when ex.Message.Contains("cmdlet", StringComparison.OrdinalIgnoreCase)
                => ErrorCode.CmdletNotAvailable,
            ErrorCategory.Operation when ex.Message.Contains("module", StringComparison.OrdinalIgnoreCase)
                => ErrorCode.ModuleNotLoaded,
            ErrorCategory.Operation when ex.Message.Contains("parameter", StringComparison.OrdinalIgnoreCase)
                => ErrorCode.InvalidParameter,
            ErrorCategory.Operation => ErrorCode.OperationNotSupported,

            ErrorCategory.Transient when ex.Message.Contains("throttl", StringComparison.OrdinalIgnoreCase)
                => ErrorCode.Throttling,
            ErrorCategory.Transient when ex.Message.Contains("timeout", StringComparison.OrdinalIgnoreCase)
                => ErrorCode.Timeout,
            ErrorCategory.Transient when ex.Message.Contains("network", StringComparison.OrdinalIgnoreCase)
                => ErrorCode.NetworkError,
            ErrorCategory.Transient => ErrorCode.ServiceUnavailable,

            ErrorCategory.Resource when ex.Message.Contains("exists", StringComparison.OrdinalIgnoreCase)
                => ErrorCode.ResourceAlreadyExists,
            ErrorCategory.Resource => ErrorCode.ResourceNotFound,

            ErrorCategory.Worker => ErrorCode.WorkerCrashed,

            _ => ErrorCode.Unknown
        };

        return new NormalizedError
        {
            Code = code,
            Message = ex.Message,
            Details = ex.InnerException?.Message,
            IsTransient = isTransient,
            RetryAfter = retryAfter.HasValue ? TimeSpan.FromSeconds(retryAfter.Value) : null,
            OriginalException = ex
        };
    }

                 
                                              
                  
    public static NormalizedError FromDto(NormalizedErrorDto dto)
    {
        return new NormalizedError
        {
            Code = dto.Code,
            Message = dto.Message,
            Details = dto.Details,
            IsTransient = dto.IsTransient,
            RetryAfter = dto.RetryAfterSeconds.HasValue
                ? TimeSpan.FromSeconds(dto.RetryAfterSeconds.Value)
                : null
        };
    }

                 
                                
                  
    public NormalizedErrorDto ToDto()
    {
        return new NormalizedErrorDto
        {
            Code = Code,
            Message = Message,
            Details = Details,
            IsTransient = IsTransient,
            RetryAfterSeconds = RetryAfter.HasValue ? (int)RetryAfter.Value.TotalSeconds : null,
            InnerException = OriginalException?.InnerException?.Message,
            StackTrace = OriginalException?.StackTrace
        };
    }

                 
                              
                  
    public static NormalizedError Create(ErrorCode code, string message, bool isTransient = false, TimeSpan? retryAfter = null)
    {
        return new NormalizedError
        {
            Code = code,
            Message = message,
            IsTransient = isTransient,
            RetryAfter = retryAfter
        };
    }
}
