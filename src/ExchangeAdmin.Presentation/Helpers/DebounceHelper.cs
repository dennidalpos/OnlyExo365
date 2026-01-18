namespace ExchangeAdmin.Presentation.Helpers;

/// <summary>
/// Helper per debounce di azioni.
/// </summary>
public sealed class DebounceHelper : IDisposable
{
    private CancellationTokenSource? _cts;
    private readonly object _lock = new();

    /// <summary>
    /// Esegue un'azione con debounce.
    /// </summary>
    /// <param name="action">Azione da eseguire.</param>
    /// <param name="delayMs">Delay in millisecondi.</param>
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
                // Debounce interrotto
            }
        }, token);
    }

    /// <summary>
    /// Esegue un'azione async con debounce.
    /// </summary>
    /// <param name="action">Azione asincrona da eseguire.</param>
    /// <param name="delayMs">Delay in millisecondi.</param>
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
                // Debounce interrotto
            }
        }, token);
    }

    /// <summary>
    /// Cancella eventuali azioni pendenti.
    /// </summary>
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
