using System.Management.Automation;
using System.Text.RegularExpressions;
using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Worker.PowerShell;

             
                                                                 
                                                                                               
              
public static partial class ErrorClassifier
{
    #region Compiled Regex Patterns

                                                                                        
    [GeneratedRegex(@"retry.+?(\d+)\s*seconds?", RegexOptions.IgnoreCase)]
    private static partial Regex ThrottlingRetryAfterRegex();

                                                                
    [GeneratedRegex(@"Retry-After:\s*(\d+)", RegexOptions.IgnoreCase)]
    private static partial Regex RetryAfterHeaderRegex();

                                                            
    [GeneratedRegex(@"back.?off.+?(\d+)", RegexOptions.IgnoreCase)]
    private static partial Regex BackoffRegex();

                                                                     
    [GeneratedRegex(@"AADSTS(\d{5,6})", RegexOptions.IgnoreCase)]
    private static partial Regex AadErrorCodeRegex();

    #endregion

    #region FullyQualifiedErrorId Known Values

                 
                                                                    
                  
    private static readonly HashSet<string> AuthenticationErrorIds = new(StringComparer.OrdinalIgnoreCase)
    {
        "AuthenticationFailed",
        "UnauthorizedAccess",
        "CredentialExpired",
        "InvalidCredential",
        "AuthenticationError",
        "TokenRequestFailed",
        "InteractiveAuthenticationRequired",
        "MsalUiRequiredException"
    };

                 
                                                              
                  
    private static readonly HashSet<string> PermissionErrorIds = new(StringComparer.OrdinalIgnoreCase)
    {
        "PermissionDenied",
        "AccessDenied",
        "InsufficientPermissions",
        "ManagementObjectNotFound,Microsoft.Exchange.Configuration.Tasks.GetMailboxPermissionTask",
        "UnauthorizedAccessException"
    };

                 
                                                                            
                  
    private static readonly HashSet<string> CmdletErrorIds = new(StringComparer.OrdinalIgnoreCase)
    {
        "CommandNotFoundException",
        "ParameterNotFound",
        "ParameterBindingException",
        "InvalidOperation",
        "NamedParameterNotFound",
        "MissingMandatoryParameter"
    };

                 
                                             
                  
    private static readonly HashSet<string> MfaRequiredAadstsCodes = new()
    {
        "50076",                                  
        "50079",                     
        "50072",                           
        "53000",                          
        "53001",                          
        "53002",                          
        "53003"                                  
    };

                 
                                                     
                  
    private static readonly HashSet<string> TokenErrorAadstsCodes = new()
    {
        "70008",                                    
        "70011",                   
        "70012",                                    
        "50173",                       
        "50078",                        
        "65001",                    
        "700016"                         
    };

    #endregion

                 
                                                                       
                  
                                                                
                                                                                 
    public static (ErrorCode Code, bool IsTransient, int? RetryAfterSeconds) Classify(ErrorRecord error)
    {
        if (error == null)
        {
            return (ErrorCode.Unknown, false, null);
        }

        var message = error.Exception?.Message ?? error.ToString();
        var categoryInfo = error.CategoryInfo;
        var fullyQualifiedErrorId = error.FullyQualifiedErrorId ?? string.Empty;
        var exceptionType = error.Exception?.GetType().Name ?? string.Empty;

                                                                                    
        var errorIdParts = fullyQualifiedErrorId.Split(',');
        var primaryErrorId = errorIdParts.Length > 0 ? errorIdParts[0].Trim() : string.Empty;

        return ClassifyInternal(message, fullyQualifiedErrorId, primaryErrorId, exceptionType, categoryInfo);
    }

                 
                                                                   
                  
                                                           
                                                                                 
    public static (ErrorCode Code, bool IsTransient, int? RetryAfterSeconds) Classify(Exception ex)
    {
        if (ex == null)
        {
            return (ErrorCode.Unknown, false, null);
        }

        var message = ex.Message;
        var exceptionType = ex.GetType().Name;

        return ClassifyInternal(message, string.Empty, string.Empty, exceptionType, null);
    }

                 
                                                                 
                  
    public static (ErrorCode Code, bool IsTransient, int? RetryAfterSeconds) ClassifyMessage(
        string message,
        string errorId,
        string exceptionType,
        ErrorCategoryInfo? categoryInfo)
    {
        var errorIdParts = (errorId ?? string.Empty).Split(',');
        var primaryErrorId = errorIdParts.Length > 0 ? errorIdParts[0].Trim() : string.Empty;

        return ClassifyInternal(message ?? string.Empty, errorId ?? string.Empty, primaryErrorId, exceptionType ?? string.Empty, categoryInfo);
    }

    private static (ErrorCode Code, bool IsTransient, int? RetryAfterSeconds) ClassifyInternal(
        string message,
        string fullErrorId,
        string primaryErrorId,
        string exceptionType,
        ErrorCategoryInfo? categoryInfo)
    {
        var messageLower = message.ToLowerInvariant();
        var errorIdLower = fullErrorId.ToLowerInvariant();

        // Check for deprecation warnings that should be ignored
        if (IsDeprecationWarning(messageLower))
        {
            // Return a special code that indicates this is just a warning, not an error
            return (ErrorCode.Unknown, false, null);
        }


        if (!string.IsNullOrEmpty(primaryErrorId))
        {
            if (AuthenticationErrorIds.Contains(primaryErrorId))
            {
                return ClassifyAuthenticationSubtype(messageLower);
            }

            if (PermissionErrorIds.Contains(primaryErrorId))
            {
                return ClassifyPermissionSubtype(messageLower);
            }

            if (CmdletErrorIds.Contains(primaryErrorId))
            {
                return ClassifyCmdletSubtype(messageLower, primaryErrorId);
            }
        }

                                                                   
        var aadResult = CheckAadErrorCodes(message);
        if (aadResult.HasValue)
        {
            return aadResult.Value;
        }

                                             
        if (categoryInfo != null)
        {
            var categoryResult = ClassifyByCategory(categoryInfo, messageLower);
            if (categoryResult.HasValue)
            {
                return categoryResult.Value;
            }
        }

                                          

                                                                     
        if (IsThrottlingError(messageLower, errorIdLower))
        {
            var retryAfter = ExtractRetryAfterSeconds(message);
            return (ErrorCode.Throttling, true, retryAfter);
        }

                                
        if (IsAuthenticationError(messageLower, errorIdLower, exceptionType))
        {
            return ClassifyAuthenticationSubtype(messageLower);
        }

                            
        if (IsPermissionError(messageLower, errorIdLower))
        {
            return ClassifyPermissionSubtype(messageLower);
        }

                                 
        if (IsTransientError(messageLower, errorIdLower, exceptionType))
        {
            if (messageLower.Contains("timeout") || messageLower.Contains("timed out"))
            {
                return (ErrorCode.Timeout, true, null);
            }

            if (messageLower.Contains("network") || messageLower.Contains("connection"))
            {
                return (ErrorCode.NetworkError, true, null);
            }

            return (ErrorCode.ServiceUnavailable, true, null);
        }

                                  
        if (IsCmdletError(messageLower, errorIdLower))
        {
            return ClassifyCmdletSubtype(messageLower, primaryErrorId);
        }

                          
        if (IsResourceError(messageLower, errorIdLower))
        {
            if (messageLower.Contains("already exists") || messageLower.Contains("duplicate"))
            {
                return (ErrorCode.ResourceAlreadyExists, false, null);
            }

            return (ErrorCode.ResourceNotFound, false, null);
        }

        return (ErrorCode.Unknown, false, null);
    }

    #region Classification Helpers

    private static (ErrorCode Code, bool IsTransient, int? RetryAfterSeconds) ClassifyAuthenticationSubtype(string messageLower)
    {
        if (messageLower.Contains("mfa") ||
            messageLower.Contains("multi-factor") ||
            messageLower.Contains("strong authentication") ||
            messageLower.Contains("two-factor"))
        {
            return (ErrorCode.MfaRequired, false, null);
        }

        if (messageLower.Contains("conditional access") ||
            messageLower.Contains("ca policy") ||
            messageLower.Contains("blocked by policy"))
        {
            return (ErrorCode.ConditionalAccessBlocked, false, null);
        }

        if (messageLower.Contains("token") && (messageLower.Contains("expired") || messageLower.Contains("invalid")))
        {
            return (ErrorCode.TokenExpired, false, null);
        }

        if (messageLower.Contains("refresh token") && (messageLower.Contains("expired") || messageLower.Contains("revoked")))
        {
            return (ErrorCode.TokenExpired, false, null);
        }

        return (ErrorCode.AuthenticationFailed, false, null);
    }

    private static (ErrorCode Code, bool IsTransient, int? RetryAfterSeconds) ClassifyPermissionSubtype(string messageLower)
    {
        if (messageLower.Contains("insufficient") || messageLower.Contains("privilege"))
        {
            return (ErrorCode.InsufficientPrivileges, false, null);
        }

        return (ErrorCode.PermissionDenied, false, null);
    }

    private static (ErrorCode Code, bool IsTransient, int? RetryAfterSeconds) ClassifyCmdletSubtype(string messageLower, string errorId)
    {
        if (messageLower.Contains("cmdlet") ||
            messageLower.Contains("command") && messageLower.Contains("not") ||
            messageLower.Contains("is not recognized") ||
            errorId.Equals("CommandNotFoundException", StringComparison.OrdinalIgnoreCase))
        {
            return (ErrorCode.CmdletNotAvailable, false, null);
        }

        if (messageLower.Contains("module"))
        {
            return (ErrorCode.ModuleNotLoaded, false, null);
        }

        if (messageLower.Contains("parameter") ||
            errorId.Contains("Parameter", StringComparison.OrdinalIgnoreCase))
        {
            return (ErrorCode.InvalidParameter, false, null);
        }

        return (ErrorCode.OperationNotSupported, false, null);
    }

    private static (ErrorCode Code, bool IsTransient, int? RetryAfterSeconds)? CheckAadErrorCodes(string message)
    {
        var aadMatch = AadErrorCodeRegex().Match(message);
        if (!aadMatch.Success)
        {
            return null;
        }

        var aadCode = aadMatch.Groups[1].Value;

        if (MfaRequiredAadstsCodes.Contains(aadCode))
        {
            return (ErrorCode.MfaRequired, false, null);
        }

        if (TokenErrorAadstsCodes.Contains(aadCode))
        {
            return (ErrorCode.TokenExpired, false, null);
        }

                                                              
        if (aadCode.StartsWith("500") || aadCode.StartsWith("700"))
        {
            return (ErrorCode.AuthenticationFailed, false, null);
        }

        return null;
    }

    private static (ErrorCode Code, bool IsTransient, int? RetryAfterSeconds)? ClassifyByCategory(
        ErrorCategoryInfo categoryInfo,
        string messageLower)
    {
        return categoryInfo.Category switch
        {
            ErrorCategory.PermissionDenied => (ErrorCode.PermissionDenied, false, null),
            ErrorCategory.SecurityError => (ErrorCode.PermissionDenied, false, null),
            ErrorCategory.AuthenticationError => ClassifyAuthenticationSubtype(messageLower),
            ErrorCategory.ResourceUnavailable => (ErrorCode.ServiceUnavailable, true, null),
            ErrorCategory.ConnectionError => (ErrorCode.NetworkError, true, null),
            ErrorCategory.OperationTimeout => (ErrorCode.Timeout, true, null),
            ErrorCategory.ObjectNotFound when messageLower.Contains("command") => (ErrorCode.CmdletNotAvailable, false, null),
            ErrorCategory.ObjectNotFound => (ErrorCode.ResourceNotFound, false, null),
            ErrorCategory.ResourceExists => (ErrorCode.ResourceAlreadyExists, false, null),
            ErrorCategory.InvalidArgument => (ErrorCode.InvalidParameter, false, null),
            ErrorCategory.InvalidOperation => (ErrorCode.OperationNotSupported, false, null),
            _ => null
        };
    }

    #endregion

    #region Pattern Detection

    private static bool IsDeprecationWarning(string message)
    {
        return message.Contains("deprecat") ||
               message.Contains("will start deprecating") ||
               message.Contains("please refer to");
    }

    private static bool IsAuthenticationError(string message, string errorId, string exceptionType)
    {
        return message.Contains("authentication") ||
               message.Contains("unauthorized") ||
               message.Contains("access denied") && message.Contains("sign in") ||
               message.Contains("login") && message.Contains("fail") ||
               message.Contains("credential") ||
               message.Contains("aadsts") ||
               errorId.Contains("authentication") ||
               exceptionType.Contains("Authentication") ||
               exceptionType.Contains("Msal");
    }

    private static bool IsPermissionError(string message, string errorId)
    {
        return message.Contains("permission") ||
               message.Contains("access denied") && !message.Contains("sign in") ||
               message.Contains("not authorized") ||
               message.Contains("forbidden") ||
               message.Contains("rbac") ||
               errorId.Contains("permission") ||
               errorId.Contains("accessdenied");
    }

    private static bool IsThrottlingError(string message, string errorId)
    {
        return message.Contains("throttl") ||
               message.Contains("rate limit") ||
               message.Contains("too many requests") ||
               message.Contains("toomanyrequests") ||
               message.Contains("429") ||
               message.Contains("back off") ||
               message.Contains("backoff") ||
               message.Contains("retry-after") ||
               errorId.Contains("throttl") ||
               errorId.Contains("ratelimit");
    }

    private static bool IsTransientError(string message, string errorId, string exceptionType)
    {
        return message.Contains("temporarily unavailable") ||
               message.Contains("service unavailable") ||
               message.Contains("503") ||
               message.Contains("502") ||
               message.Contains("504") ||
               message.Contains("500") && message.Contains("internal server") ||
               message.Contains("timeout") ||
               message.Contains("timed out") ||
               message.Contains("operation timed out") ||
               message.Contains("network") && message.Contains("error") ||
               message.Contains("connection") && (message.Contains("fail") || message.Contains("refused") || message.Contains("reset") || message.Contains("closed")) ||
               message.Contains("transient") ||
               exceptionType.Contains("TimeoutException") ||
               exceptionType.Contains("WebException") ||
               exceptionType.Contains("HttpRequestException") ||
               exceptionType.Contains("SocketException");
    }

    private static bool IsCmdletError(string message, string errorId)
    {
        return message.Contains("cmdlet") ||
               message.Contains("command not found") ||
               message.Contains("is not recognized") ||
               message.Contains("not recognized as") ||
               message.Contains("module") && message.Contains("not") ||
               message.Contains("parameter") && (message.Contains("not found") || message.Contains("invalid") || message.Contains("cannot bind")) ||
               errorId.Contains("commandnotfound") ||
               errorId.Contains("parameternotfound");
    }

    private static bool IsResourceError(string message, string errorId)
    {
        return message.Contains("not found") && !message.Contains("command") && !message.Contains("cmdlet") ||
               message.Contains("doesn't exist") ||
               message.Contains("does not exist") ||
               message.Contains("couldn't be found") ||
               message.Contains("already exists") ||
               message.Contains("duplicate") ||
               errorId.Contains("objectnotfound");
    }

    #endregion

    #region Retry-After Extraction

                 
                                                               
                                                                                           
                  
                                                          
                                                                     
    private static int? ExtractRetryAfterSeconds(string message)
    {
        if (string.IsNullOrEmpty(message))
        {
            return null;
        }

                                             
        var headerMatch = RetryAfterHeaderRegex().Match(message);
        if (headerMatch.Success && int.TryParse(headerMatch.Groups[1].Value, out var headerSeconds))
        {
            return headerSeconds;
        }

                                              
        var retryMatch = ThrottlingRetryAfterRegex().Match(message);
        if (retryMatch.Success && int.TryParse(retryMatch.Groups[1].Value, out var retrySeconds))
        {
            return retrySeconds;
        }

                              
        var backoffMatch = BackoffRegex().Match(message);
        if (backoffMatch.Success && int.TryParse(backoffMatch.Groups[1].Value, out var backoffSeconds))
        {
            return backoffSeconds;
        }

        return null;
    }

    #endregion
}
