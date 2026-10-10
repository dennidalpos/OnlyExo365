using System.Windows;
using System.Windows.Threading;

namespace OnlyExo365.Shell.Helpers;

public sealed class DebounceHelper : IDisposable
{
    private readonly Dispatcher _dispatcher;
    private readonly DispatcherTimer _timer;
    private Action? _pendingAction;
    private bool _disposed;

    public DebounceHelper()
    {
        _dispatcher = Application.Current?.Dispatcher ?? Dispatcher.CurrentDispatcher;
        _timer = new DispatcherTimer(DispatcherPriority.Background, _dispatcher);
        _timer.Tick += OnTick;
        _dispatcher.ShutdownStarted += OnShutdownStarted;
    }

    public void Debounce(Action action, int delayMs = 300)
    {
        ArgumentNullException.ThrowIfNull(action);
        ArgumentOutOfRangeException.ThrowIfNegative(delayMs);
        OnDispatcher(() =>
        {
            if (_disposed) return;
            _timer.Stop();
            _pendingAction = action;
            _timer.Interval = TimeSpan.FromMilliseconds(delayMs);
            _timer.Start();
        });
    }

    public void Cancel() => OnDispatcher(CancelCore);

    public void Dispose() => OnDispatcher(() =>
    {
        if (_disposed) return;
        _disposed = true;
        CancelCore();
        _timer.Tick -= OnTick;
        _dispatcher.ShutdownStarted -= OnShutdownStarted;
    });

    private void OnTick(object? sender, EventArgs e)
    {
        var action = _pendingAction;
        CancelCore();
        action?.Invoke();
    }

    private void CancelCore()
    {
        _timer.Stop();
        _pendingAction = null;
    }

    private void OnShutdownStarted(object? sender, EventArgs e) => Dispose();

    private void OnDispatcher(Action action)
    {
        if (_dispatcher.CheckAccess()) action();
        else if (!_dispatcher.HasShutdownStarted) _dispatcher.Invoke(action);
    }
}
