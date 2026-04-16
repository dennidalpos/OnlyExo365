using ExchangeAdmin.Domain.Errors;

namespace ExchangeAdmin.Domain.Results;

             
                                                       
              
public class Result
{
    public bool IsSuccess { get; protected init; }
    public bool IsFailure => !IsSuccess;
    public NormalizedError? Error { get; protected init; }
    public bool WasCancelled { get; protected init; }
    public string? CorrelationId { get; protected init; }

    protected Result() { }

    public static Result Success(string? correlationId = null) => new()
    {
        IsSuccess = true,
        CorrelationId = correlationId
    };

    public static Result Failure(NormalizedError error, string? correlationId = null) => new()
    {
        IsSuccess = false,
        Error = error,
        CorrelationId = correlationId
    };

    public static Result Cancelled(string? correlationId = null) => new()
    {
        IsSuccess = false,
        WasCancelled = true,
        CorrelationId = correlationId
    };

    public static Result FromException(Exception ex, string? correlationId = null)
    {
        if (ex is OperationCanceledException)
            return Cancelled(correlationId);

        return Failure(NormalizedError.FromException(ex), correlationId);
    }
}

             
                                                     
              
public class Result<T> : Result
{
    public T? Value { get; private init; }

    private Result() { }

    public static Result<T> Success(T value, string? correlationId = null) => new()
    {
        IsSuccess = true,
        Value = value,
        CorrelationId = correlationId
    };

    public new static Result<T> Failure(NormalizedError error, string? correlationId = null) => new()
    {
        IsSuccess = false,
        Error = error,
        CorrelationId = correlationId
    };

    public new static Result<T> Cancelled(string? correlationId = null) => new()
    {
        IsSuccess = false,
        WasCancelled = true,
        CorrelationId = correlationId
    };

    public new static Result<T> FromException(Exception ex, string? correlationId = null)
    {
        if (ex is OperationCanceledException)
            return Cancelled(correlationId);

        return Failure(NormalizedError.FromException(ex), correlationId);
    }

                 
                                            
                  
    public Result<TNew> Map<TNew>(Func<T, TNew> mapper)
    {
        if (IsSuccess && Value != null)
            return Result<TNew>.Success(mapper(Value), CorrelationId);

        if (WasCancelled)
            return Result<TNew>.Cancelled(CorrelationId);

        return Result<TNew>.Failure(Error!, CorrelationId);
    }

                 
                                  
                  
    public Result<T> OnSuccess(Action<T> action)
    {
        if (IsSuccess && Value != null)
            action(Value);
        return this;
    }

                 
                                    
                  
    public Result<T> OnFailure(Action<NormalizedError> action)
    {
        if (IsFailure && Error != null)
            action(Error);
        return this;
    }
}
