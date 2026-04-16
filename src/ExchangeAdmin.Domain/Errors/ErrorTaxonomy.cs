namespace ExchangeAdmin.Domain.Errors;

             
                                            
              
public enum ErrorCategory
{
                 
                                                                          
                  
    Authentication,

                 
                                                                 
                  
    Permission,

                 
                                                                              
                  
    Operation,

                 
                                                                      
                  
    Transient,

                 
                                                     
                  
    Resource,

                 
                          
                  
    Worker,

                 
                                 
                  
    Unknown
}

             
                                                        
              
public static class ErrorTaxonomy
{
    private static readonly Dictionary<string, (ErrorCategory Category, bool IsTransient)> KnownPatterns = new(StringComparer.OrdinalIgnoreCase)
    {
                                
        ["AADSTS50076"] = (ErrorCategory.Authentication, false),                
        ["AADSTS53003"] = (ErrorCategory.Authentication, false),                              
        ["AADSTS50058"] = (ErrorCategory.Authentication, false),                         
        ["AADSTS700016"] = (ErrorCategory.Authentication, false),                 
        ["AADSTS65001"] = (ErrorCategory.Authentication, false),                    
        ["token has expired"] = (ErrorCategory.Authentication, false),
        ["token is expired"] = (ErrorCategory.Authentication, false),
        ["refresh token has expired"] = (ErrorCategory.Authentication, false),

                            
        ["Access is denied"] = (ErrorCategory.Permission, false),
        ["AccessDenied"] = (ErrorCategory.Permission, false),
        ["Insufficient permissions"] = (ErrorCategory.Permission, false),
        ["doesn't have the required permissions"] = (ErrorCategory.Permission, false),
        ["ManagementObjectNotFoundException"] = (ErrorCategory.Permission, false),

                                  
        ["is not recognized as the name of a cmdlet"] = (ErrorCategory.Operation, false),
        ["The term"] = (ErrorCategory.Operation, false),
        ["A parameter cannot be found"] = (ErrorCategory.Operation, false),
        ["Cannot validate argument on parameter"] = (ErrorCategory.Operation, false),
        ["Cannot bind parameter"] = (ErrorCategory.Operation, false),
        ["ParameterBindingException"] = (ErrorCategory.Operation, false),

                           
        ["throttl"] = (ErrorCategory.Transient, true),
        ["too many requests"] = (ErrorCategory.Transient, true),
        ["429"] = (ErrorCategory.Transient, true),
        ["503"] = (ErrorCategory.Transient, true),
        ["service unavailable"] = (ErrorCategory.Transient, true),
        ["temporarily unavailable"] = (ErrorCategory.Transient, true),
        ["connection was forcibly closed"] = (ErrorCategory.Transient, true),
        ["timeout"] = (ErrorCategory.Transient, true),
        ["network"] = (ErrorCategory.Transient, true),

                          
        ["couldn't be found"] = (ErrorCategory.Resource, false),
        ["does not exist"] = (ErrorCategory.Resource, false),
        ["was not found"] = (ErrorCategory.Resource, false),
        ["already exists"] = (ErrorCategory.Resource, false)
    };

                 
                                                     
                  
    public static (ErrorCategory Category, bool IsTransient) Classify(string? errorMessage, string? exceptionType = null)
    {
        if (string.IsNullOrWhiteSpace(errorMessage))
            return (ErrorCategory.Unknown, false);

                             
        foreach (var pattern in KnownPatterns)
        {
            if (errorMessage.Contains(pattern.Key, StringComparison.OrdinalIgnoreCase))
            {
                return pattern.Value;
            }
        }

                                                  
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

                 
                                                                
                  
    public static int? ExtractRetryAfter(string? errorMessage)
    {
        if (string.IsNullOrWhiteSpace(errorMessage))
            return null;

                                                                    
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
