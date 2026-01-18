using System.Diagnostics;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Errors;

namespace ExchangeAdmin.Domain.Resilience;

/// <summary>
/// Configurazione policy di retry.
/// </summary>
public class RetryPolicyOptions
{
    /// <summary>
    /// Numero massimo di tentativi (incluso il primo).
    /// Default: 3.
    /// </summary>
    public int MaxRetries { get; init; } = 3;

    /// <summary>
    /// Delay base tra tentativi (ms).
    /// Default: 1000ms.
    /// </summary>
    public int BaseDelayMs { get; init; } = 1000;

    /// <summary>
    /// Delay massimo tra tentativi (ms). Il backoff non supererà mai questo valore.
    /// Default: 30000ms (30 secondi).
    /// </summary>
    public int MaxDelayMs { get; init; } = 30000;

    /// <summary>
    /// Fattore di backoff esponenziale.
    /// Default: 2.0 (delay raddoppia ad ogni tentativo).
    /// </summary>
    public double BackoffFactor { get; init; } = 2.0;

    /// <summary>
    /// Usa decorrelated jitter invece di jitter semplice.
    /// Decorrelated jitter produce distribuzioni più uniformi e previene thundering herd.
    /// Default: true.
    /// </summary>
    public bool UseDecorrelatedJitter { get; init; } = true;

    /// <summary>
    /// Jitter massimo per jitter semplice (frazione del delay, 0-1).
    /// Ignorato se UseDecorrelatedJitter è true.
    /// Default: 0.2 (±20%).
    /// </summary>
    public double MaxJitter { get; init; } = 0.2;
}

/// <summary>
/// Risultato di una operazione con retry.
/// </summary>
/// <typeparam name="T">Tipo del risultato.</typeparam>
public class RetryResult<T>
{
    /// <summary>True se l'operazione ha avuto successo.</summary>
    public bool Success { get; init; }

    /// <summary>Risultato dell'operazione (null se fallita).</summary>
    public T? Result { get; init; }

    /// <summary>Errore normalizzato (null se successo).</summary>
    public NormalizedError? Error { get; init; }

    /// <summary>Numero di tentativi effettuati.</summary>
    public int Attempts { get; init; }

    /// <summary>Durata totale inclusi tutti i retry e delay.</summary>
    public TimeSpan TotalDuration { get; init; }

    /// <summary>True se l'operazione è stata cancellata.</summary>
    public bool WasCancelled { get; init; }
}

/// <summary>
/// Lista di ErrorCode che non devono essere ritentati.
/// Questi errori sono permanenti e ritentarli non cambierà il risultato.
/// </summary>
public static class NonRetryableErrors
{
    /// <summary>
    /// ErrorCodes che non devono mai essere ritentati.
    /// </summary>
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

    /// <summary>
    /// Verifica se un ErrorCode è non-retryable.
    /// </summary>
    public static bool IsNonRetryable(ErrorCode code) => Codes.Contains(code);
}

/// <summary>
/// Policy di retry con exponential backoff e decorrelated jitter.
/// Implementa AWS-style decorrelated jitter per distribuzioni uniformi.
/// </summary>
/// <remarks>
/// <para>
/// Decorrelated jitter (formula): sleep = min(cap, random_between(base, sleep * 3))
/// Questo produce delay più uniformi e previene il "thundering herd" problem.
/// </para>
/// <para>
/// La policy non riprova errori permanenti come PermissionDenied o InvalidParameter.
/// Solo errori transient (throttling, network, timeout) vengono ritentati.
/// </para>
/// </remarks>
public class RetryPolicy
{
    private readonly RetryPolicyOptions _options;
    private readonly ITimeProvider _timeProvider;
    private readonly Random _random;

    /// <summary>
    /// Crea una nuova retry policy.
    /// </summary>
    /// <param name="options">Opzioni di configurazione (null = default).</param>
    /// <param name="timeProvider">Provider del tempo per testabilità (null = sistema).</param>
    /// <param name="randomSeed">Seed per random (null = non-deterministic).</param>
    public RetryPolicy(RetryPolicyOptions? options = null, ITimeProvider? timeProvider = null, int? randomSeed = null)
    {
        _options = options ?? new RetryPolicyOptions();
        _timeProvider = timeProvider ?? SystemTimeProvider.Instance;
        _random = randomSeed.HasValue ? new Random(randomSeed.Value) : new Random();
    }

    /// <summary>
    /// Esegue un'operazione con retry policy.
    /// </summary>
    /// <typeparam name="T">Tipo del risultato.</typeparam>
    /// <param name="operation">Operazione da eseguire.</param>
    /// <param name="cancellationToken">Token di cancellazione.</param>
    /// <returns>Risultato con metadata sui tentativi.</returns>
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

                // Non riprovare errori permanenti
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

                // Solo retry per errori transient
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

                // Verifica se abbiamo esaurito i tentativi
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

                // Calcola delay
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

    /// <summary>
    /// Esegue un'operazione senza valore di ritorno con retry policy.
    /// </summary>
    /// <param name="operation">Operazione da eseguire.</param>
    /// <param name="cancellationToken">Token di cancellazione.</param>
    /// <returns>Risultato con metadata sui tentativi.</returns>
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

    /// <summary>
    /// Calcola delay con exponential backoff e jitter.
    /// </summary>
    /// <param name="attempt">Numero del tentativo corrente (1-based).</param>
    /// <param name="retryAfter">Retry-after suggerito dal server (opzionale).</param>
    /// <param name="previousDelay">Delay del tentativo precedente (per decorrelated jitter).</param>
    /// <returns>Delay da attendere prima del prossimo tentativo.</returns>
    private TimeSpan CalculateDelay(int attempt, TimeSpan? retryAfter, ref double previousDelay)
    {
        double delayMs;

        // Se il server ha specificato retry-after, usalo come minimo
        var serverRetryAfterMs = retryAfter?.TotalMilliseconds ?? 0;

        if (_options.UseDecorrelatedJitter)
        {
            // Decorrelated jitter (AWS style)
            // Formula: sleep = min(cap, random_between(base, sleep * 3))
            var minDelay = _options.BaseDelayMs;
            var maxDelay = Math.Min(_options.MaxDelayMs, previousDelay * 3);

            delayMs = minDelay + _random.NextDouble() * (maxDelay - minDelay);
            previousDelay = delayMs;
        }
        else
        {
            // Exponential backoff con jitter semplice
            var baseDelay = _options.BaseDelayMs * Math.Pow(_options.BackoffFactor, attempt - 1);
            baseDelay = Math.Min(baseDelay, _options.MaxDelayMs);

            var jitter = baseDelay * _options.MaxJitter * (_random.NextDouble() * 2 - 1);
            delayMs = baseDelay + jitter;
        }

        // Usa il maggiore tra delay calcolato e retry-after del server
        delayMs = Math.Max(delayMs, serverRetryAfterMs);

        // Clamp al massimo e assicura positivo
        delayMs = Math.Min(delayMs, _options.MaxDelayMs);
        delayMs = Math.Max(delayMs, 0);

        return TimeSpan.FromMilliseconds(delayMs);
    }
}
