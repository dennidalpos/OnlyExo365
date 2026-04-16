using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using ExchangeAdmin.Presentation.Localization;
using ExchangeAdmin.Presentation.Services;
using ExchangeAdmin.Presentation.Text;

namespace ExchangeAdmin.Presentation.ViewModels;

public abstract class ViewModelBase : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;

    protected ViewModelBase()
    {
        // All ViewModels are application-lifetime singletons; no memory leak concern.
        LocalizationService.Instance.CultureChanged += OnCultureChanged;
    }

    private void OnCultureChanged(object? sender, EventArgs e)
    {
        RunOnUiThread(() => OnPropertyChanged(string.Empty));
    }

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

    protected void RunOnUiThread(Action action)
    {
        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        if (dispatcher == null)
        {
            action();
            return;
        }

        if (dispatcher.CheckAccess())
        {
            action();
        }
        else
        {
            dispatcher.Invoke(action);
        }
    }

    protected async Task RunOnUiThreadAsync(Action action)
    {
        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        if (dispatcher == null)
        {
            action();
            return;
        }

        if (dispatcher.CheckAccess())
        {
            action();
        }
        else
        {
            await dispatcher.InvokeAsync(action);
        }
    }

    protected static string FormatProgressPercent(double percent)
    {
        var normalized = Math.Clamp((int)Math.Round(percent), 0, 100);
        return $"{normalized}%";
    }

    protected static string? FormatProgressCount(int? currentItem, int? totalItems, string? fallbackCompletedLabel = null)
    {
        if (currentItem.HasValue && totalItems.HasValue && totalItems.Value > 0)
        {
            var detected = Math.Clamp(currentItem.Value, 0, totalItems.Value);
            var missing = Math.Max(totalItems.Value - detected, 0);
            return Loc.GetFormat("Progress.LoadedRemaining", detected, missing);
        }

        if (currentItem.HasValue && currentItem.Value > 0)
        {
            return fallbackCompletedLabel != null
                ? Loc.GetFormat("Progress.LoadedWithLabel", currentItem.Value, fallbackCompletedLabel)
                : Loc.GetFormat("Progress.LoadedCount", currentItem.Value);
        }

        return null;
    }

    protected bool ConfirmMutation(string operation, string target, string? impact = null, string? title = null)
    {
        var confirmed = false;
        RunOnUiThread(() =>
        {
            confirmed = ErrorDialogService.ShowConfirmation(
                title ?? UserMessageCatalog.ConfirmOperationTitle,
                UserMessageCatalog.FormatMutationConfirmation(operation, target, impact));
        });

        return confirmed;
    }
}
