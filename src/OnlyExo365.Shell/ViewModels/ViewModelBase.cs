using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using OnlyExo365.Shell.Localization;
using OnlyExo365.Shell.Services;
using OnlyExo365.Shell.Text;

namespace OnlyExo365.Shell.ViewModels;

public abstract class ViewModelBase : INotifyPropertyChanged
{
    private readonly List<IDisposable> _ownedResources = [];
    private bool _disposed;
    public event PropertyChangedEventHandler? PropertyChanged;

    protected T Own<T>(T resource) where T : IDisposable
    {
        _ownedResources.Add(resource);
        return resource;
    }

    protected OnlyExo365.Shell.Helpers.DebounceHelper CreateDebounceHelper()
        => Own(new OnlyExo365.Shell.Helpers.DebounceHelper());

    public virtual void ReleaseUiResources()
    {
        if (_disposed) return;
        _disposed = true;
        LocalizationService.Instance.CultureChanged -= OnCultureChanged;
        foreach (var resource in _ownedResources) resource.Dispose();
        _ownedResources.Clear();
    }

    protected ViewModelBase()
    {
        // ViewModels live until shutdown.
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

