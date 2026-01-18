namespace ExchangeAdmin.Domain.Errors;

/// <summary>
/// Categoria di errore per classificazione.
/// </summary>
public enum ErrorCategory
{
    /// <summary>
    /// Errori di autenticazione (MFA, Conditional Access, token scaduto).
    /// </summary>
    Authentication,

    /// <summary>
    /// Errori di permesso (ruoli insufficienti, accesso negato).
    /// </summary>
    Permission,

    /// <summary>
    /// Errori cmdlet/operazione (cmdlet non disponibile, parametri invalidi).
    /// </summary>
    Operation,

    /// <summary>
    /// Errori transient (throttling, servizio non disponibile, rete).
    /// </summary>
    Transient,

    /// <summary>
    /// Errori risorsa (non trovata, già esistente).
    /// </summary>
    Resource,

    /// <summary>
    /// Errori worker/IPC.
    /// </summary>
    Worker,

    /// <summary>
    /// Errori non categorizzati.
    /// </summary>
    Unknown
}

/// <summary>
/// Classificatore di errori PowerShell Exchange Online.
/// </summary>
public static class ErrorTaxonomy
{
    private static readonly Dictionary<string, (ErrorCategory Category, bool IsTransient)> KnownPatterns = new(StringComparer.OrdinalIgnoreCase)
    {
        // Authentication errors
        ["AADSTS50076"] = (ErrorCategory.Authentication, false), // MFA required
        ["AADSTS53003"] = (ErrorCategory.Authentication, false), // Conditional Access blocked
        ["AADSTS50058"] = (ErrorCategory.Authentication, false), // Silent sign-in failed
        ["AADSTS700016"] = (ErrorCategory.Authentication, false), // App not found
        ["AADSTS65001"] = (ErrorCategory.Authentication, false), // Consent required
        ["token has expired"] = (ErrorCategory.Authentication, false),
        ["token is expired"] = (ErrorCategory.Authentication, false),
        ["refresh token has expired"] = (ErrorCategory.Authentication, false),

        // Permission errors
        ["Access is denied"] = (ErrorCategory.Permission, false),
        ["AccessDenied"] = (ErrorCategory.Permission, false),
        ["Insufficient permissions"] = (ErrorCategory.Permission, false),
        ["doesn't have the required permissions"] = (ErrorCategory.Permission, false),
        ["ManagementObjectNotFoundException"] = (ErrorCategory.Permission, false),

        // Cmdlet/Operation errors
        ["is not recognized as the name of a cmdlet"] = (ErrorCategory.Operation, false),
        ["The term"] = (ErrorCategory.Operation, false),
        ["A parameter cannot be found"] = (ErrorCategory.Operation, false),
        ["Cannot validate argument on parameter"] = (ErrorCategory.Operation, false),
        ["Cannot bind parameter"] = (ErrorCategory.Operation, false),
        ["ParameterBindingException"] = (ErrorCategory.Operation, false),

        // Transient errors
        ["throttl"] = (ErrorCategory.Transient, true),
        ["too many requests"] = (ErrorCategory.Transient, true),
        ["429"] = (ErrorCategory.Transient, true),
        ["503"] = (ErrorCategory.Transient, true),
        ["service unavailable"] = (ErrorCategory.Transient, true),
        ["temporarily unavailable"] = (ErrorCategory.Transient, true),
        ["connection was forcibly closed"] = (ErrorCategory.Transient, true),
        ["timeout"] = (ErrorCategory.Transient, true),
        ["network"] = (ErrorCategory.Transient, true),

        // Resource errors
        ["couldn't be found"] = (ErrorCategory.Resource, false),
        ["does not exist"] = (ErrorCategory.Resource, false),
        ["was not found"] = (ErrorCategory.Resource, false),
        ["already exists"] = (ErrorCategory.Resource, false)
    };

    /// <summary>
    /// Classifica un errore basandosi sul messaggio.
    /// </summary>
    public static (ErrorCategory Category, bool IsTransient) Classify(string? errorMessage, string? exceptionType = null)
    {
        if (string.IsNullOrWhiteSpace(errorMessage))
            return (ErrorCategory.Unknown, false);

        // Cerca pattern noti
        foreach (var pattern in KnownPatterns)
        {
            if (errorMessage.Contains(pattern.Key, StringComparison.OrdinalIgnoreCase))
            {
                return pattern.Value;
            }
        }

        // Controlla tipo eccezione se disponibile
        if (!string.IsNullOrWhiteSpace(exceptionType))
        {
            if (exceptionType.Contains("Authentication", StringComparison.OrdinalIgnoreCase))
                return (ErrorCategory.Authentication, false);

            if (exceptionType.Contains("Unauthorized", StringComparison.OrdinalIgnoreCase))
                return (ErrorCategory.Permission, false);

            if (exceptionType.Contains("Timeout", StringComparison.OrdinalIgnoreCase))
                return (ErrorCategory.Transient, true);
        }

        return (ErrorCategory.Unknown, false);
    }

    /// <summary>
    /// Estrae retry-after in secondi se presente nel messaggio.
    /// </summary>
    public static int? ExtractRetryAfter(string? errorMessage)
    {
        if (string.IsNullOrWhiteSpace(errorMessage))
            return null;

        // Pattern comuni: "retry after X seconds", "Retry-After: X"
        var patterns = new[]
        {
            @"retry.?after[:\s]+(\d+)",
            @"wait[:\s]+(\d+)\s*second",
            @"(\d+)\s*second.*retry"
        };

        foreach (var pattern in patterns)
        {
            var match = System.Text.RegularExpressions.Regex.Match(
                errorMessage,
                pattern,
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            if (match.Success && int.TryParse(match.Groups[1].Value, out var seconds))
            {
                return seconds;
            }
        }

        return null;
    }
}
