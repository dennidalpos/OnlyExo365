using System.Diagnostics;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Errors;

namespace ExchangeAdmin.Domain.Resilience;

             
                                   
              
public class RetryPolicyOptions
{
                 
                                                       
                   
                  
    public int MaxRetries { get; init; } = 3;

                 
                                      
                        
                  
    public int BaseDelayMs { get; init; } = 1000;

                 
                                                                                     
                                      
                  
    public int MaxDelayMs { get; init; } = 30000;

                 
                                        
                                                         
                  
    public double BackoffFactor { get; init; } = 2.0;

                 
                                                          
                                                                                           
                      
                  
    public bool UseDecorrelatedJitter { get; init; } = true;

                 
                                                                     
                                                  
                             
                  
    public double MaxJitter { get; init; } = 0.2;
}

             
                                          
              
                                                       
public class RetryResult<T>
{
                                                                  
    public bool Success { get; init; }

                                                                       
    public T? Result { get; init; }

                                                                  
    public NormalizedError? Error { get; init; }

                                                          
    public int Attempts { get; init; }

                                                                       
    public TimeSpan TotalDuration { get; init; }

                                                                    
    public bool WasCancelled { get; init; }
}

             
                                                       
                                                                          
              
public static class NonRetryableErrors
{
                 
                                                       
                  
    public static readonly HashSet<ErrorCode> Codes = new()
    {
        ErrorCode.PermissionDenied,
        ErrorCode.InsufficientPrivileges,
        ErrorCode.InvalidParameter,
        ErrorCode.CmdletNotAvailable,
        ErrorCode.ModuleNotLoaded,
        ErrorCode.OperationNotSupported,
        ErrorCode.ResourceAlreadyExists,
        ErrorCode.AuthenticationFailed,
        ErrorCode.ConditionalAccessBlocked
    };

                 
                                                  
                  
    public static bool IsNonRetryable(ErrorCode code) => Codes.Contains(code);
}

             
                                                                  
                                                                        
              
             
          
                                                                                    
                                                                               
           
          
                                                                                     
                                                                           
           
              
public class RetryPolicy
{
    private readonly RetryPolicyOptions _options;
    private readonly ITimeProvider _timeProvider;
    private readonly Random _random;

                 
                                    
                  
                                                                                 
                                                                                                
                                                                                    
    public RetryPolicy(RetryPolicyOptions? options = null, ITimeProvider? timeProvider = null, int? randomSeed = null)
    {
        _options = options ?? new RetryPolicyOptions();
        _timeProvider = timeProvider ?? SystemTimeProvider.Instance;
        _random = randomSeed.HasValue ? new Random(randomSeed.Value) : new Random();
    }

                 
                                              
                  
                                                           
                                                               
                                                                       
                                                                
    public async Task<RetryResult<T>> ExecuteAsync<T>(
        Func<CancellationToken, Task<T>> operation,
        CancellationToken cancellationToken = default)
    {
        var startTime = _timeProvider.UtcNow;
        var attempt = 0;
        NormalizedError? lastError = null;
        var previousDelay = (double)_options.BaseDelayMs;

        while (attempt < _options.MaxRetries)
        {
            attempt++;

            if (cancellationToken.IsCancellationRequested)
            {
                return new RetryResult<T>
                {
                    Success = false,
                    WasCancelled = true,
                    Attempts = attempt,
                    TotalDuration = _timeProvider.UtcNow - startTime
                };
            }

            try
            {
                Debug.WriteLine($"[RetryPolicy] Attempt {attempt}/{_options.MaxRetries}");

                var result = await operation(cancellationToken).ConfigureAwait(false);

                Debug.WriteLine($"[RetryPolicy] Success on attempt {attempt}");

                return new RetryResult<T>
                {
                    Success = true,
                    Result = result,
                    Attempts = attempt,
                    TotalDuration = _timeProvider.UtcNow - startTime
                };
            }
            catch (OperationCanceledException)
            {
                Debug.WriteLine($"[RetryPolicy] Cancelled on attempt {attempt}");

                return new RetryResult<T>
                {
                    Success = false,
                    WasCancelled = true,
                    Attempts = attempt,
                    TotalDuration = _timeProvider.UtcNow - startTime
                };
            }
            catch (Exception ex)
            {
                lastError = NormalizedError.FromException(ex);

                Debug.WriteLine($"[RetryPolicy] Attempt {attempt} failed: {lastError.Code} - {lastError.Message}");

                                                  
                if (NonRetryableErrors.IsNonRetryable(lastError.Code))
                {
                    Debug.WriteLine($"[RetryPolicy] Error {lastError.Code} is non-retryable, stopping");

                    return new RetryResult<T>
                    {
                        Success = false,
                        Error = lastError,
                        Attempts = attempt,
                        TotalDuration = _timeProvider.UtcNow - startTime
                    };
                }

                                                  
                if (!lastError.IsTransient)
                {
                    Debug.WriteLine($"[RetryPolicy] Error is not transient, stopping");

                    return new RetryResult<T>
                    {
                        Success = false,
                        Error = lastError,
                        Attempts = attempt,
                        TotalDuration = _timeProvider.UtcNow - startTime
                    };
                }

                                                           
                if (attempt >= _options.MaxRetries)
                {
                    Debug.WriteLine($"[RetryPolicy] Max retries ({_options.MaxRetries}) exhausted");

                    return new RetryResult<T>
                    {
                        Success = false,
                        Error = lastError,
                        Attempts = attempt,
                        TotalDuration = _timeProvider.UtcNow - startTime
                    };
                }

                                
                var delay = CalculateDelay(attempt, lastError.RetryAfter, ref previousDelay);

                Debug.WriteLine($"[RetryPolicy] Waiting {delay.TotalMilliseconds:F0}ms before retry {attempt + 1}");

                await Task.Delay(delay, cancellationToken).ConfigureAwait(false);
            }
        }

        return new RetryResult<T>
        {
            Success = false,
            Error = lastError,
            Attempts = attempt,
            TotalDuration = _timeProvider.UtcNow - startTime
        };
    }

                 
                                                                      
                  
                                                               
                                                                       
                                                                
    public async Task<RetryResult<bool>> ExecuteAsync(
        Func<CancellationToken, Task> operation,
        CancellationToken cancellationToken = default)
    {
        return await ExecuteAsync(async ct =>
        {
            await operation(ct).ConfigureAwait(false);
            return true;
        }, cancellationToken).ConfigureAwait(false);
    }

                 
                                                       
                  
                                                                              
                                                                                      
                                                                                                     
                                                                           
    private TimeSpan CalculateDelay(int attempt, TimeSpan? retryAfter, ref double previousDelay)
    {
        double delayMs;

                                                                     
        var serverRetryAfterMs = retryAfter?.TotalMilliseconds ?? 0;

        if (_options.UseDecorrelatedJitter)
        {
                                              
                                                                         
            var minDelay = _options.BaseDelayMs;
            var maxDelay = Math.Min(_options.MaxDelayMs, previousDelay * 3);

            delayMs = minDelay + _random.NextDouble() * (maxDelay - minDelay);
            previousDelay = delayMs;
        }
        else
        {
                                                      
            var baseDelay = _options.BaseDelayMs * Math.Pow(_options.BackoffFactor, attempt - 1);
            baseDelay = Math.Min(baseDelay, _options.MaxDelayMs);

            var jitter = baseDelay * _options.MaxJitter * (_random.NextDouble() * 2 - 1);
            delayMs = baseDelay + jitter;
        }

                                                                       
        delayMs = Math.Max(delayMs, serverRetryAfterMs);

                                               
        delayMs = Math.Min(delayMs, _options.MaxDelayMs);
        delayMs = Math.Max(delayMs, 0);

        return TimeSpan.FromMilliseconds(delayMs);
    }
}
