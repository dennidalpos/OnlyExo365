namespace ExchangeAdmin.Presentation.Helpers;

             
                                  
              
public sealed class DebounceHelper : IDisposable
{
    private CancellationTokenSource? _cts;
    private readonly object _lock = new();

                 
                                      
                  
                                                        
                                                            
    public void Debounce(Action action, int delayMs = 300)
    {
        lock (_lock)
        {
            _cts?.Cancel();
            _cts?.Dispose();
            _cts = new CancellationTokenSource();
        }

        var token = _cts.Token;

        Task.Run(async () =>
        {
            try
            {
                await Task.Delay(delayMs, token);
                if (!token.IsCancellationRequested)
                {
                    action();
                }
            }
            catch (OperationCanceledException)
            {
                                      
            }
        }, token);
    }

                 
                                            
                  
                                                                  
                                                            
    public void DebounceAsync(Func<CancellationToken, Task> action, int delayMs = 300)
    {
        lock (_lock)
        {
            _cts?.Cancel();
            _cts?.Dispose();
            _cts = new CancellationTokenSource();
        }

        var token = _cts.Token;

        Task.Run(async () =>
        {
            try
            {
                await Task.Delay(delayMs, token);
                if (!token.IsCancellationRequested)
                {
                    await action(token);
                }
            }
            catch (OperationCanceledException)
            {
                                      
            }
        }, token);
    }

                 
                                           
                  
    public void Cancel()
    {
        lock (_lock)
        {
            _cts?.Cancel();
        }
    }

    public void Dispose()
    {
        _cts?.Cancel();
        _cts?.Dispose();
    }
}
