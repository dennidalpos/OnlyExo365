using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;

namespace ExchangeAdmin.Presentation.ViewModels;

/// <summary>
/// Base class per ViewModels con INotifyPropertyChanged.
/// </summary>
public abstract class ViewModelBase : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
            return false;

        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    /// <summary>
    /// Esegue un'azione sul dispatcher UI.
    /// </summary>
    protected void RunOnUiThread(Action action)
    {
        if (System.Windows.Application.Current?.Dispatcher?.CheckAccess() == true)
        {
            action();
        }
        else
        {
            System.Windows.Application.Current?.Dispatcher?.Invoke(action);
        }
    }

    /// <summary>
    /// Esegue un'azione asincrona sul dispatcher UI.
    /// </summary>
    protected async Task RunOnUiThreadAsync(Action action)
    {
        if (System.Windows.Application.Current?.Dispatcher?.CheckAccess() == true)
        {
            action();
        }
        else
        {
            var dispatcher = System.Windows.Application.Current?.Dispatcher;
            if (dispatcher != null)
            {
                await dispatcher.InvokeAsync(action);
            }
        }
    }
}
