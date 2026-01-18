using System.Diagnostics;

namespace ExchangeAdmin.Domain.Resilience;

/// <summary>
/// Stato del circuit breaker.
/// </summary>
public enum CircuitState
{
    /// <summary>Normale - lascia passare tutte le richieste.</summary>
    Closed,

    /// <summary>Aperto - blocca le richieste, restituisce errore immediato.</summary>
    Open,

    /// <summary>Semi-aperto - consente una richiesta di prova per verificare recovery.</summary>
    HalfOpen
}

/// <summary>
/// Configurazione circuit breaker.
/// </summary>
public class CircuitBreakerOptions
{
    /// <summary>
    /// Numero di fallimenti consecutivi per aprire il circuito.
    /// Default: 5.
    /// </summary>
    public int FailureThreshold { get; init; } = 5;

    /// <summary>
    /// Durata stato Open prima di passare a HalfOpen.
    /// Default: 30 secondi.
    /// </summary>
    public TimeSpan OpenDuration { get; init; } = TimeSpan.FromSeconds(30);

    /// <summary>
    /// Numero di successi consecutivi in HalfOpen per chiudere il circuito.
    /// Default: 2.
    /// </summary>
    public int SuccessThresholdInHalfOpen { get; init; } = 2;
}

/// <summary>
/// Astrazione del tempo per testabilità deterministica.
/// </summary>
public interface ITimeProvider
{
    /// <summary>Ottiene il tempo corrente UTC.</summary>
    DateTime UtcNow { get; }
}

/// <summary>
/// Implementazione default che usa il tempo di sistema.
/// </summary>
public sealed class SystemTimeProvider : ITimeProvider
{
    /// <summary>Istanza singleton.</summary>
    public static readonly SystemTimeProvider Instance = new();

    private SystemTimeProvider() { }

    /// <inheritdoc />
    public DateTime UtcNow => DateTime.UtcNow;
}

/// <summary>
/// Circuit breaker thread-safe per proteggere da fallimenti ripetuti.
/// Implementa il pattern Circuit Breaker con tre stati: Closed, Open, HalfOpen.
/// </summary>
/// <remarks>
/// <para>
/// Quando il circuito è <see cref="CircuitState.Closed"/>, tutte le richieste passano.
/// Dopo <see cref="CircuitBreakerOptions.FailureThreshold"/> fallimenti consecutivi,
/// il circuito si apre.
/// </para>
/// <para>
/// Quando il circuito è <see cref="CircuitState.Open"/>, tutte le richieste falliscono
/// immediatamente con <see cref="CircuitBreakerOpenException"/>. Dopo
/// <see cref="CircuitBreakerOptions.OpenDuration"/>, il circuito passa a HalfOpen.
/// </para>
/// <para>
/// Quando il circuito è <see cref="CircuitState.HalfOpen"/>, una richiesta di prova
/// è consentita. Se ha successo per <see cref="CircuitBreakerOptions.SuccessThresholdInHalfOpen"/>
/// volte consecutive, il circuito si chiude. Se fallisce, torna in Open.
/// </para>
/// </remarks>
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

    /// <summary>
    /// Evento scatenato quando lo stato del circuito cambia.
    /// </summary>
    public event EventHandler<CircuitStateChangedEventArgs>? StateChanged;

    /// <summary>
    /// Crea un nuovo circuit breaker con opzioni e time provider custom.
    /// </summary>
    /// <param name="options">Opzioni di configurazione (null = default).</param>
    /// <param name="timeProvider">Provider del tempo (null = sistema).</param>
    public CircuitBreaker(CircuitBreakerOptions? options = null, ITimeProvider? timeProvider = null)
    {
        _options = options ?? new CircuitBreakerOptions();
        _timeProvider = timeProvider ?? SystemTimeProvider.Instance;
        _lastStateChangeTime = _timeProvider.UtcNow;
    }

    /// <summary>
    /// Stato corrente del circuito.
    /// L'accesso a questa proprietà può causare transizione automatica da Open a HalfOpen.
    /// </summary>
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

    /// <summary>
    /// Numero di fallimenti consecutivi corrente.
    /// </summary>
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

    /// <summary>
    /// Timestamp dell'ultimo cambio di stato.
    /// </summary>
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

    /// <summary>
    /// Tempo rimanente prima che il circuito passi da Open a HalfOpen.
    /// Restituisce TimeSpan.Zero se non in stato Open.
    /// </summary>
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

    /// <summary>
    /// Verifica se il circuito consente l'esecuzione di una richiesta.
    /// </summary>
    /// <returns>True se la richiesta può procedere, false se bloccata.</returns>
    public bool CanExecute()
    {
        var state = State; // Trigger auto-transizione
        return state != CircuitState.Open;
    }

    /// <summary>
    /// Registra un'operazione completata con successo.
    /// In stato HalfOpen, contribuisce al conteggio per chiudere il circuito.
    /// In stato Closed, resetta il contatore fallimenti.
    /// </summary>
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
                // Reset contatore fallimenti su successo
                if (_failureCount > 0)
                {
                    Debug.WriteLine($"[CircuitBreaker] Success - resetting failure count from {_failureCount}");
                    _failureCount = 0;
                }
            }
        }
    }

    /// <summary>
    /// Registra un fallimento (solo per errori transient).
    /// Non chiamare per errori permanenti (PermissionDenied, InvalidParameter, etc.).
    /// </summary>
    public void RecordFailure()
    {
        lock (_lock)
        {
            _lastFailureTime = _timeProvider.UtcNow;

            if (_state == CircuitState.HalfOpen)
            {
                // Torna immediatamente in Open - recovery fallito
                Debug.WriteLine("[CircuitBreaker] Failure in HalfOpen - reopening circuit");
                TransitionTo(CircuitState.Open, "Recovery failed");
                _openedAt = _timeProvider.UtcNow;
                return;
            }

            _failureCount++;
            Debug.WriteLine($"[CircuitBreaker] Failure recorded ({_failureCount}/{_options.FailureThreshold})");

            if (_failureCount >= _options.FailureThreshold)
            {
                // Apri il circuito
                TransitionTo(CircuitState.Open, $"Threshold reached ({_failureCount} failures)");
                _openedAt = _timeProvider.UtcNow;
            }
        }
    }

    /// <summary>
    /// Reset forzato del circuit breaker allo stato Closed.
    /// Usare con cautela - normalmente il circuito si auto-gestisce.
    /// </summary>
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

    /// <summary>
    /// Esegue un'operazione con protezione circuit breaker.
    /// </summary>
    /// <typeparam name="T">Tipo del risultato.</typeparam>
    /// <param name="operation">Operazione da eseguire.</param>
    /// <param name="cancellationToken">Token di cancellazione.</param>
    /// <returns>Risultato dell'operazione.</returns>
    /// <exception cref="CircuitBreakerOpenException">Se il circuito è aperto.</exception>
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
            // Cancellazione non conta come fallimento
            throw;
        }
        catch (Exception)
        {
            RecordFailure();
            throw;
        }
    }

    /// <summary>
    /// Esegue un'operazione senza valore di ritorno con protezione circuit breaker.
    /// </summary>
    /// <param name="operation">Operazione da eseguire.</param>
    /// <param name="cancellationToken">Token di cancellazione.</param>
    /// <exception cref="CircuitBreakerOpenException">Se il circuito è aperto.</exception>
    public async Task ExecuteAsync(Func<CancellationToken, Task> operation, CancellationToken cancellationToken = default)
    {
        await ExecuteAsync(async ct =>
        {
            await operation(ct).ConfigureAwait(false);
            return true;
        }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Ottiene informazioni diagnostiche sullo stato corrente.
    /// </summary>
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
        // Chiamato già dentro lock
        if (_state == CircuitState.Open &&
            _timeProvider.UtcNow - _openedAt >= _options.OpenDuration)
        {
            TransitionTo(CircuitState.HalfOpen, "Open duration elapsed");
            _successCountInHalfOpen = 0;
        }
    }

    private void TransitionTo(CircuitState newState, string reason)
    {
        // Chiamato già dentro lock
        var previousState = _state;
        _state = newState;
        _lastStateChangeTime = _timeProvider.UtcNow;

        Debug.WriteLine($"[CircuitBreaker] State: {previousState} -> {newState} ({reason})");
        OnStateChanged(previousState, newState, reason);
    }

    private void OnStateChanged(CircuitState previousState, CircuitState newState, string reason)
    {
        // Fire event fuori dal lock per evitare deadlock
        var handler = StateChanged;
        if (handler != null)
        {
            // Queue sul thread pool per non bloccare
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

/// <summary>
/// Argomenti per l'evento di cambio stato del circuit breaker.
/// </summary>
public class CircuitStateChangedEventArgs : EventArgs
{
    /// <summary>Stato precedente.</summary>
    public CircuitState PreviousState { get; }

    /// <summary>Nuovo stato.</summary>
    public CircuitState NewState { get; }

    /// <summary>Motivo della transizione.</summary>
    public string Reason { get; }

    /// <summary>Timestamp della transizione.</summary>
    public DateTime Timestamp { get; }

    /// <summary>
    /// Crea nuovi argomenti evento.
    /// </summary>
    public CircuitStateChangedEventArgs(CircuitState previousState, CircuitState newState, string reason)
    {
        PreviousState = previousState;
        NewState = newState;
        Reason = reason;
        Timestamp = DateTime.UtcNow;
    }
}

/// <summary>
/// Informazioni diagnostiche sul circuit breaker.
/// </summary>
public class CircuitBreakerDiagnostics
{
    /// <summary>Stato corrente.</summary>
    public CircuitState State { get; init; }

    /// <summary>Numero di fallimenti consecutivi.</summary>
    public int FailureCount { get; init; }

    /// <summary>Numero di successi in HalfOpen.</summary>
    public int SuccessCountInHalfOpen { get; init; }

    /// <summary>Timestamp ultimo fallimento.</summary>
    public DateTime LastFailureTime { get; init; }

    /// <summary>Timestamp apertura circuito.</summary>
    public DateTime OpenedAt { get; init; }

    /// <summary>Timestamp ultimo cambio stato.</summary>
    public DateTime LastStateChangeTime { get; init; }

    /// <summary>Tempo rimanente in stato Open.</summary>
    public TimeSpan RemainingOpenTime { get; init; }

    /// <summary>Opzioni di configurazione.</summary>
    public CircuitBreakerOptions Options { get; init; } = new();

    /// <inheritdoc />
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

/// <summary>
/// Eccezione lanciata quando il circuit breaker è aperto.
/// </summary>
public class CircuitBreakerOpenException : Exception
{
    /// <summary>
    /// Tempo rimanente prima che il circuito tenti recovery.
    /// </summary>
    public TimeSpan RemainingOpenTime { get; }

    /// <summary>
    /// Crea una nuova eccezione circuit breaker open.
    /// </summary>
    /// <param name="remainingOpenTime">Tempo rimanente.</param>
    public CircuitBreakerOpenException(TimeSpan remainingOpenTime)
        : base($"Circuit breaker is open. Will retry in {remainingOpenTime.TotalSeconds:F0} seconds.")
    {
        RemainingOpenTime = remainingOpenTime > TimeSpan.Zero ? remainingOpenTime : TimeSpan.Zero;
    }
}
