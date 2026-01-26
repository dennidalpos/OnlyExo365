using System.Diagnostics;

namespace ExchangeAdmin.Domain.Resilience;

             
                              
              
public enum CircuitState
{
                                                                       
    Closed,

                                                                                      
    Open,

                                                                                                 
    HalfOpen
}

             
                                   
              
public class CircuitBreakerOptions
{
                 
                                                                
                   
                  
    public int FailureThreshold { get; init; } = 5;

                 
                                                      
                            
                  
    public TimeSpan OpenDuration { get; init; } = TimeSpan.FromSeconds(30);

                 
                                                                            
                   
                  
    public int SuccessThresholdInHalfOpen { get; init; } = 2;
}

             
                                                         
              
public interface ITimeProvider
{
                                                         
    DateTime UtcNow { get; }
}

             
                                                        
              
public sealed class SystemTimeProvider : ITimeProvider
{
                                             
    public static readonly SystemTimeProvider Instance = new();

    private SystemTimeProvider() { }

                      
    public DateTime UtcNow => DateTime.UtcNow;
}

             
                                                                      
                                                                                
              
             
          
                                                                                        
                                                                                     
                        
           
          
                                                                                        
                                                                      
                                                                                   
           
          
                                                                                     
                                                                                                    
                                                                         
           
              
public class CircuitBreaker
{
    private readonly CircuitBreakerOptions _options;
    private readonly ITimeProvider _timeProvider;
    private readonly object _lock = new();

    private CircuitState _state = CircuitState.Closed;
    private int _failureCount;
    private int _successCountInHalfOpen;
    private DateTime _lastFailureTime;
    private DateTime _openedAt;
    private DateTime _lastStateChangeTime;

                 
                                                             
                  
    public event EventHandler<CircuitStateChangedEventArgs>? StateChanged;

                 
                                                                         
                  
                                                                                 
                                                                               
    public CircuitBreaker(CircuitBreakerOptions? options = null, ITimeProvider? timeProvider = null)
    {
        _options = options ?? new CircuitBreakerOptions();
        _timeProvider = timeProvider ?? SystemTimeProvider.Instance;
        _lastStateChangeTime = _timeProvider.UtcNow;
    }

                 
                                    
                                                                                             
                  
    public CircuitState State
    {
        get
        {
            lock (_lock)
            {
                TryTransitionToHalfOpen();
                return _state;
            }
        }
    }

                 
                                                  
                  
    public int FailureCount
    {
        get
        {
            lock (_lock)
            {
                return _failureCount;
            }
        }
    }

                 
                                              
                  
    public DateTime LastStateChangeTime
    {
        get
        {
            lock (_lock)
            {
                return _lastStateChangeTime;
            }
        }
    }

                 
                                                                       
                                                       
                  
    public TimeSpan RemainingOpenTime
    {
        get
        {
            lock (_lock)
            {
                if (_state != CircuitState.Open)
                {
                    return TimeSpan.Zero;
                }

                var elapsed = _timeProvider.UtcNow - _openedAt;
                var remaining = _options.OpenDuration - elapsed;
                return remaining > TimeSpan.Zero ? remaining : TimeSpan.Zero;
            }
        }
    }

                 
                                                                       
                  
                                                                                  
    public bool CanExecute()
    {
        var state = State;                            
        return state != CircuitState.Open;
    }

                 
                                                       
                                                                              
                                                         
                  
    public void RecordSuccess()
    {
        lock (_lock)
        {
            if (_state == CircuitState.HalfOpen)
            {
                _successCountInHalfOpen++;
                Debug.WriteLine($"[CircuitBreaker] Success in HalfOpen ({_successCountInHalfOpen}/{_options.SuccessThresholdInHalfOpen})");

                if (_successCountInHalfOpen >= _options.SuccessThresholdInHalfOpen)
                {
                    TransitionTo(CircuitState.Closed, "Recovery successful");
                    _failureCount = 0;
                }
            }
            else if (_state == CircuitState.Closed)
            {
                                                         
                if (_failureCount > 0)
                {
                    Debug.WriteLine($"[CircuitBreaker] Success - resetting failure count from {_failureCount}");
                    _failureCount = 0;
                }
            }
        }
    }

                 
                                                           
                                                                                      
                  
    public void RecordFailure()
    {
        lock (_lock)
        {
            _lastFailureTime = _timeProvider.UtcNow;

            if (_state == CircuitState.HalfOpen)
            {
                                                                  
                Debug.WriteLine("[CircuitBreaker] Failure in HalfOpen - reopening circuit");
                TransitionTo(CircuitState.Open, "Recovery failed");
                _openedAt = _timeProvider.UtcNow;
                return;
            }

            _failureCount++;
            Debug.WriteLine($"[CircuitBreaker] Failure recorded ({_failureCount}/{_options.FailureThreshold})");

            if (_failureCount >= _options.FailureThreshold)
            {
                                   
                TransitionTo(CircuitState.Open, $"Threshold reached ({_failureCount} failures)");
                _openedAt = _timeProvider.UtcNow;
            }
        }
    }

                 
                                                            
                                                                     
                  
    public void Reset()
    {
        lock (_lock)
        {
            var previousState = _state;
            _state = CircuitState.Closed;
            _failureCount = 0;
            _successCountInHalfOpen = 0;
            _lastStateChangeTime = _timeProvider.UtcNow;

            if (previousState != CircuitState.Closed)
            {
                Debug.WriteLine($"[CircuitBreaker] Manually reset from {previousState}");
                OnStateChanged(previousState, CircuitState.Closed, "Manual reset");
            }
        }
    }

                 
                                                            
                  
                                                           
                                                               
                                                                       
                                                     
                                                                                           
    public async Task<T> ExecuteAsync<T>(Func<CancellationToken, Task<T>> operation, CancellationToken cancellationToken = default)
    {
        if (!CanExecute())
        {
            var remaining = RemainingOpenTime;
            Debug.WriteLine($"[CircuitBreaker] Blocked - circuit open, retry in {remaining.TotalSeconds:F1}s");
            throw new CircuitBreakerOpenException(remaining);
        }

        try
        {
            var result = await operation(cancellationToken).ConfigureAwait(false);
            RecordSuccess();
            return result;
        }
        catch (OperationCanceledException)
        {
                                                      
            throw;
        }
        catch (Exception)
        {
            RecordFailure();
            throw;
        }
    }

                 
                                                                                    
                  
                                                               
                                                                       
                                                                                           
    public async Task ExecuteAsync(Func<CancellationToken, Task> operation, CancellationToken cancellationToken = default)
    {
        await ExecuteAsync(async ct =>
        {
            await operation(ct).ConfigureAwait(false);
            return true;
        }, cancellationToken).ConfigureAwait(false);
    }

                 
                                                               
                  
    public CircuitBreakerDiagnostics GetDiagnostics()
    {
        lock (_lock)
        {
            TryTransitionToHalfOpen();

            return new CircuitBreakerDiagnostics
            {
                State = _state,
                FailureCount = _failureCount,
                SuccessCountInHalfOpen = _successCountInHalfOpen,
                LastFailureTime = _lastFailureTime,
                OpenedAt = _openedAt,
                LastStateChangeTime = _lastStateChangeTime,
                RemainingOpenTime = _state == CircuitState.Open
                    ? _options.OpenDuration - (_timeProvider.UtcNow - _openedAt)
                    : TimeSpan.Zero,
                Options = _options
            };
        }
    }

    private void TryTransitionToHalfOpen()
    {
                                    
        if (_state == CircuitState.Open &&
            _timeProvider.UtcNow - _openedAt >= _options.OpenDuration)
        {
            TransitionTo(CircuitState.HalfOpen, "Open duration elapsed");
            _successCountInHalfOpen = 0;
        }
    }

    private void TransitionTo(CircuitState newState, string reason)
    {
                                    
        var previousState = _state;
        _state = newState;
        _lastStateChangeTime = _timeProvider.UtcNow;

        Debug.WriteLine($"[CircuitBreaker] State: {previousState} -> {newState} ({reason})");
        OnStateChanged(previousState, newState, reason);
    }

    private void OnStateChanged(CircuitState previousState, CircuitState newState, string reason)
    {
                                                         
        var handler = StateChanged;
        if (handler != null)
        {
                                                     
            _ = Task.Run(() =>
            {
                try
                {
                    handler(this, new CircuitStateChangedEventArgs(previousState, newState, reason));
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[CircuitBreaker] StateChanged handler error: {ex.Message}");
                }
            });
        }
    }
}

             
                                                               
              
public class CircuitStateChangedEventArgs : EventArgs
{
                                            
    public CircuitState PreviousState { get; }

                                       
    public CircuitState NewState { get; }

                                                    
    public string Reason { get; }

                                                       
    public DateTime Timestamp { get; }

                 
                                    
                  
    public CircuitStateChangedEventArgs(CircuitState previousState, CircuitState newState, string reason)
    {
        PreviousState = previousState;
        NewState = newState;
        Reason = reason;
        Timestamp = DateTime.UtcNow;
    }
}

             
                                                  
              
public class CircuitBreakerDiagnostics
{
                                          
    public CircuitState State { get; init; }

                                                            
    public int FailureCount { get; init; }

                                                          
    public int SuccessCountInHalfOpen { get; init; }

                                                       
    public DateTime LastFailureTime { get; init; }

                                                       
    public DateTime OpenedAt { get; init; }

                                                         
    public DateTime LastStateChangeTime { get; init; }

                                                         
    public TimeSpan RemainingOpenTime { get; init; }

                                                     
    public CircuitBreakerOptions Options { get; init; } = new();

                      
    public override string ToString()
    {
        return State switch
        {
            CircuitState.Open => $"Open (retry in {RemainingOpenTime.TotalSeconds:F0}s, {FailureCount} failures)",
            CircuitState.HalfOpen => $"HalfOpen ({SuccessCountInHalfOpen}/{Options.SuccessThresholdInHalfOpen} successes)",
            _ => $"Closed ({FailureCount}/{Options.FailureThreshold} failures)"
        };
    }
}

             
                                                           
              
public class CircuitBreakerOpenException : Exception
{
                 
                                                             
                  
    public TimeSpan RemainingOpenTime { get; }

                 
                                                      
                  
                                                                
    public CircuitBreakerOpenException(TimeSpan remainingOpenTime)
        : base($"Circuit breaker is open. Will retry in {remainingOpenTime.TotalSeconds:F0} seconds.")
    {
        RemainingOpenTime = remainingOpenTime > TimeSpan.Zero ? remainingOpenTime : TimeSpan.Zero;
    }
}
