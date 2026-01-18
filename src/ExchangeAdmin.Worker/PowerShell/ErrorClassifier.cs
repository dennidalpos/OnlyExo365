using System.Management.Automation;
using System.Text.RegularExpressions;
using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Worker.PowerShell;

/// <summary>
/// Classifica errori PowerShell/EXO in error codes normalizzati.
/// Implementa parsing robusto di FullyQualifiedErrorId, CategoryInfo, e pattern di throttling.
/// </summary>
public static partial class ErrorClassifier
{
    #region Compiled Regex Patterns

    /// <summary>Pattern per estrazione retry-after da messaggi di throttling.</summary>
    [GeneratedRegex(@"retry.+?(\d+)\s*seconds?", RegexOptions.IgnoreCase)]
    private static partial Regex ThrottlingRetryAfterRegex();

    /// <summary>Pattern per Retry-After header style.</summary>
    [GeneratedRegex(@"Retry-After:\s*(\d+)", RegexOptions.IgnoreCase)]
    private static partial Regex RetryAfterHeaderRegex();

    /// <summary>Pattern per backoff instructions.</summary>
    [GeneratedRegex(@"back.?off.+?(\d+)", RegexOptions.IgnoreCase)]
    private static partial Regex BackoffRegex();

    /// <summary>Pattern per Azure AD error codes (AADSTS).</summary>
    [GeneratedRegex(@"AADSTS(\d{5,6})", RegexOptions.IgnoreCase)]
    private static partial Regex AadErrorCodeRegex();

    #endregion

    #region FullyQualifiedErrorId Known Values

    /// <summary>
    /// FullyQualifiedErrorId che indicano errori di autenticazione.
    /// </summary>
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

    /// <summary>
    /// FullyQualifiedErrorId che indicano errori di permessi.
    /// </summary>
    private static readonly HashSet<string> PermissionErrorIds = new(StringComparer.OrdinalIgnoreCase)
    {
        "PermissionDenied",
        "AccessDenied",
        "InsufficientPermissions",
        "ManagementObjectNotFound,Microsoft.Exchange.Configuration.Tasks.GetMailboxPermissionTask",
        "UnauthorizedAccessException"
    };

    /// <summary>
    /// FullyQualifiedErrorId che indicano cmdlet/parametri non disponibili.
    /// </summary>
    private static readonly HashSet<string> CmdletErrorIds = new(StringComparer.OrdinalIgnoreCase)
    {
        "CommandNotFoundException",
        "ParameterNotFound",
        "ParameterBindingException",
        "InvalidOperation",
        "NamedParameterNotFound",
        "MissingMandatoryParameter"
    };

    /// <summary>
    /// AADSTS error codes per MFA richiesta.
    /// </summary>
    private static readonly HashSet<string> MfaRequiredAadstsCodes = new()
    {
        "50076", // Strong authentication required
        "50079", // User must use MFA
        "50072", // MFA enrollment required
        "53000", // Conditional access MFA
        "53001", // Conditional access MFA
        "53002", // App cannot be accessed
        "53003"  // Blocked by conditional access
    };

    /// <summary>
    /// AADSTS error codes per token expired/invalid.
    /// </summary>
    private static readonly HashSet<string> TokenErrorAadstsCodes = new()
    {
        "70008", // Expired or revoked refresh token
        "70011", // Malformed scope
        "70012", // MsAccount authentication failure
        "50173", // Fresh auth required
        "50078", // Fresh claim required
        "65001", // Consent required
        "700016" // Application not found
    };

    #endregion

    /// <summary>
    /// Classifica un ErrorRecord PowerShell in un errore normalizzato.
    /// </summary>
    /// <param name="error">ErrorRecord da classificare.</param>
    /// <returns>Tupla con (ErrorCode, IsTransient, RetryAfterSeconds).</returns>
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

        // Estrai la parte principale dell'ErrorId (prima della virgola se presente)
        var errorIdParts = fullyQualifiedErrorId.Split(',');
        var primaryErrorId = errorIdParts.Length > 0 ? errorIdParts[0].Trim() : string.Empty;

        return ClassifyInternal(message, fullyQualifiedErrorId, primaryErrorId, exceptionType, categoryInfo);
    }

    /// <summary>
    /// Classifica un'eccezione generica in un errore normalizzato.
    /// </summary>
    /// <param name="ex">Eccezione da classificare.</param>
    /// <returns>Tupla con (ErrorCode, IsTransient, RetryAfterSeconds).</returns>
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

    /// <summary>
    /// Classifica un messaggio di errore con metadati opzionali.
    /// </summary>
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

        // 1. Check FullyQualifiedErrorId first (most specific)
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

        // 2. Check for AADSTS codes in message (Azure AD specific)
        var aadResult = CheckAadErrorCodes(message);
        if (aadResult.HasValue)
        {
            return aadResult.Value;
        }

        // 3. Check CategoryInfo if available
        if (categoryInfo != null)
        {
            var categoryResult = ClassifyByCategory(categoryInfo, messageLower);
            if (categoryResult.HasValue)
            {
                return categoryResult.Value;
            }
        }

        // 4. Pattern-based classification

        // Throttling (transient) - check first as it's common in EXO
        if (IsThrottlingError(messageLower, errorIdLower))
        {
            var retryAfter = ExtractRetryAfterSeconds(message);
            return (ErrorCode.Throttling, true, retryAfter);
        }

        // Authentication errors
        if (IsAuthenticationError(messageLower, errorIdLower, exceptionType))
        {
            return ClassifyAuthenticationSubtype(messageLower);
        }

        // Permission errors
        if (IsPermissionError(messageLower, errorIdLower))
        {
            return ClassifyPermissionSubtype(messageLower);
        }

        // Other transient errors
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

        // Cmdlet/operation errors
        if (IsCmdletError(messageLower, errorIdLower))
        {
            return ClassifyCmdletSubtype(messageLower, primaryErrorId);
        }

        // Resource errors
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

        // AADSTS50001, 500XX sono spesso errori di app/config
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

    /// <summary>
    /// Estrae il valore retry-after da un messaggio di errore.
    /// Supporta vari formati: "retry after X seconds", "Retry-After: X", "back off for X".
    /// </summary>
    /// <param name="message">Messaggio di errore.</param>
    /// <returns>Secondi da attendere, null se non trovato.</returns>
    private static int? ExtractRetryAfterSeconds(string message)
    {
        if (string.IsNullOrEmpty(message))
        {
            return null;
        }

        // Try Retry-After header style first
        var headerMatch = RetryAfterHeaderRegex().Match(message);
        if (headerMatch.Success && int.TryParse(headerMatch.Groups[1].Value, out var headerSeconds))
        {
            return headerSeconds;
        }

        // Try "retry after X seconds" pattern
        var retryMatch = ThrottlingRetryAfterRegex().Match(message);
        if (retryMatch.Success && int.TryParse(retryMatch.Groups[1].Value, out var retrySeconds))
        {
            return retrySeconds;
        }

        // Try backoff pattern
        var backoffMatch = BackoffRegex().Match(message);
        if (backoffMatch.Success && int.TryParse(backoffMatch.Groups[1].Value, out var backoffSeconds))
        {
            return backoffSeconds;
        }

        return null;
    }

    #endregion
}
