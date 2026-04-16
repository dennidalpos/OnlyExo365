using System.Windows.Input;

namespace ExchangeAdmin.Presentation.Helpers;

public class RelayCommand : ICommand
{
    private readonly Action<object?> _execute;
    private readonly Func<object?, bool>? _canExecute;

    public RelayCommand(Action<object?> execute, Func<object?, bool>? canExecute = null)
    {
        _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        _canExecute = canExecute;
    }

    public RelayCommand(Action execute, Func<bool>? canExecute = null)
        : this(_ => execute(), canExecute != null ? _ => canExecute() : null)
    {
    }

    public event EventHandler? CanExecuteChanged
    {
        add => CommandManager.RequerySuggested += value;
        remove => CommandManager.RequerySuggested -= value;
    }

    public bool CanExecute(object? parameter) => _canExecute?.Invoke(parameter) ?? true;

    public void Execute(object? parameter) => _execute(parameter);

    public void RaiseCanExecuteChanged()
    {
        CommandManager.InvalidateRequerySuggested();
    }
}

public class RelayCommand<T> : ICommand
{
    private readonly Action<T?> _execute;
    private readonly Func<T?, bool>? _canExecute;

    public RelayCommand(Action<T?> execute, Func<T?, bool>? canExecute = null)
    {
        _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        _canExecute = canExecute;
    }

    public event EventHandler? CanExecuteChanged
    {
        add => CommandManager.RequerySuggested += value;
        remove => CommandManager.RequerySuggested -= value;
    }

    public bool CanExecute(object? parameter)
    {
        if (parameter is T typedParam)
            return _canExecute?.Invoke(typedParam) ?? true;
        if (parameter == null)
            return _canExecute?.Invoke(default) ?? true;
        return false;
    }

    public void Execute(object? parameter)
    {
        if (parameter is T typedParam)
            _execute(typedParam);
        else if (parameter == null)
            _execute(default);
    }

    public void RaiseCanExecuteChanged()
    {
        CommandManager.InvalidateRequerySuggested();
    }
}

public class AsyncRelayCommand : ICommand
{
    private readonly Func<CancellationToken, Task> _execute;
    private readonly Func<bool>? _canExecute;
    private CancellationTokenSource? _cts;
    private bool _isExecuting;

    public bool IsExecuting => _isExecuting;

    public AsyncRelayCommand(Func<CancellationToken, Task> execute, Func<bool>? canExecute = null)
    {
        _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        _canExecute = canExecute;
    }

    public AsyncRelayCommand(Func<Task> execute, Func<bool>? canExecute = null)
        : this(_ => execute(), canExecute)
    {
    }

    public event EventHandler? CanExecuteChanged
    {
        add => CommandManager.RequerySuggested += value;
        remove => CommandManager.RequerySuggested -= value;
    }

    public bool CanExecute(object? parameter)
    {
        return !_isExecuting && (_canExecute?.Invoke() ?? true);
    }

    public async void Execute(object? parameter)
    {
        if (_isExecuting) return;

        _isExecuting = true;
        _cts = new CancellationTokenSource();
        RaiseCanExecuteChanged();

        try
        {
            await _execute(_cts.Token);
        }
        catch (OperationCanceledException)
        {
            // Intentionally ignored.
        }
        finally
        {
            _isExecuting = false;
            _cts?.Dispose();
            _cts = null;
            RaiseCanExecuteChanged();
        }
    }

    public void Cancel()
    {
        _cts?.Cancel();
    }

    public void RaiseCanExecuteChanged()
    {
        CommandManager.InvalidateRequerySuggested();
    }
}

public class AsyncRelayCommand<T> : ICommand
{
    private readonly Func<T?, CancellationToken, Task> _execute;
    private readonly Func<T?, bool>? _canExecute;
    private CancellationTokenSource? _cts;
    private bool _isExecuting;

    public bool IsExecuting => _isExecuting;

    public AsyncRelayCommand(Func<T?, CancellationToken, Task> execute, Func<T?, bool>? canExecute = null)
    {
        _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        _canExecute = canExecute;
    }

    public AsyncRelayCommand(Func<T?, Task> execute, Func<T?, bool>? canExecute = null)
        : this((p, _) => execute(p), canExecute)
    {
    }

    public event EventHandler? CanExecuteChanged
    {
        add => CommandManager.RequerySuggested += value;
        remove => CommandManager.RequerySuggested -= value;
    }

    public bool CanExecute(object? parameter)
    {
        if (_isExecuting) return false;

        if (parameter is T typedParam)
            return _canExecute?.Invoke(typedParam) ?? true;
        if (parameter == null)
            return _canExecute?.Invoke(default) ?? true;
        return false;
    }

    public async void Execute(object? parameter)
    {
        if (_isExecuting) return;

        T? typedParam = default;
        if (parameter is T p)
            typedParam = p;

        _isExecuting = true;
        _cts = new CancellationTokenSource();
        RaiseCanExecuteChanged();

        try
        {
            await _execute(typedParam, _cts.Token);
        }
        catch (OperationCanceledException)
        {
            // Intentionally ignored.
        }
        finally
        {
            _isExecuting = false;
            _cts?.Dispose();
            _cts = null;
            RaiseCanExecuteChanged();
        }
    }

    public void Cancel()
    {
        _cts?.Cancel();
    }

    public void RaiseCanExecuteChanged()
    {
        CommandManager.InvalidateRequerySuggested();
    }
}
