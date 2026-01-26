using ExchangeAdmin.Domain.Errors;

namespace ExchangeAdmin.Domain.Results;

             
                                                       
              
public class Result
{
    public bool IsSuccess { get; protected init; }
    public bool IsFailure => !IsSuccess;
    public NormalizedError? Error { get; protected init; }
    public bool WasCancelled { get; protected init; }

    protected Result() { }

    public static Result Success() => new() { IsSuccess = true };

    public static Result Failure(NormalizedError error) => new()
    {
        IsSuccess = false,
        Error = error
    };

    public static Result Cancelled() => new()
    {
        IsSuccess = false,
        WasCancelled = true
    };

    public static Result FromException(Exception ex)
    {
        if (ex is OperationCanceledException)
            return Cancelled();

        return Failure(NormalizedError.FromException(ex));
    }
}

             
                                                     
              
public class Result<T> : Result
{
    public T? Value { get; private init; }

    private Result() { }

    public static Result<T> Success(T value) => new()
    {
        IsSuccess = true,
        Value = value
    };

    public new static Result<T> Failure(NormalizedError error) => new()
    {
        IsSuccess = false,
        Error = error
    };

    public new static Result<T> Cancelled() => new()
    {
        IsSuccess = false,
        WasCancelled = true
    };

    public new static Result<T> FromException(Exception ex)
    {
        if (ex is OperationCanceledException)
            return Cancelled();

        return Failure(NormalizedError.FromException(ex));
    }

                 
                                            
                  
    public Result<TNew> Map<TNew>(Func<T, TNew> mapper)
    {
        if (IsSuccess && Value != null)
            return Result<TNew>.Success(mapper(Value));

        if (WasCancelled)
            return Result<TNew>.Cancelled();

        return Result<TNew>.Failure(Error!);
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
