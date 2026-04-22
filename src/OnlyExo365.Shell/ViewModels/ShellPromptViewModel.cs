using System.Windows.Input;
using System.Windows.Threading;
using OnlyExo365.Shell.Helpers;

namespace OnlyExo365.Shell.ViewModels;

public enum ShellPromptKind
{
    Information,
    Warning,
    Confirmation
}

public sealed class ShellPromptViewModel : ViewModelBase
{
    private bool _isOpen;
    private string _title = string.Empty;
    private string _message = string.Empty;
    private ShellPromptKind _kind = ShellPromptKind.Information;
    private TaskCompletionSource<bool>? _pendingConfirmation;

    public ShellPromptViewModel()
    {
        ConfirmCommand = new RelayCommand(Confirm, () => IsOpen && Kind == ShellPromptKind.Confirmation);
        CancelCommand = new RelayCommand(Cancel, () => IsOpen);
    }

    public bool IsOpen
    {
        get => _isOpen;
        private set
        {
            if (SetProperty(ref _isOpen, value))
            {
                OnPropertyChanged(nameof(IsConfirmation));
                OnPropertyChanged(nameof(IsWarning));
                OnPropertyChanged(nameof(IsInformation));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string Title
    {
        get => _title;
        private set => SetProperty(ref _title, value);
    }

    public string Message
    {
        get => _message;
        private set => SetProperty(ref _message, value);
    }

    public ShellPromptKind Kind
    {
        get => _kind;
        private set
        {
            if (SetProperty(ref _kind, value))
            {
                OnPropertyChanged(nameof(IsConfirmation));
                OnPropertyChanged(nameof(IsWarning));
                OnPropertyChanged(nameof(IsInformation));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool IsConfirmation => IsOpen && Kind == ShellPromptKind.Confirmation;
    public bool IsWarning => IsOpen && Kind == ShellPromptKind.Warning;
    public bool IsInformation => IsOpen && Kind == ShellPromptKind.Information;

    public ICommand ConfirmCommand { get; }
    public ICommand CancelCommand { get; }

    public void ShowInformation(string title, string message)
        => ShowMessage(title, message, ShellPromptKind.Information);

    public void ShowWarning(string title, string message)
        => ShowMessage(title, message, ShellPromptKind.Warning);

    public bool ShowConfirmationBlocking(string title, string message)
    {
        if (System.Windows.Application.Current?.Dispatcher is not Dispatcher dispatcher)
        {
            return false;
        }

        if (!dispatcher.CheckAccess())
        {
            return dispatcher.Invoke(() => ShowConfirmationBlocking(title, message));
        }

        ResolvePendingConfirmation(false);

        var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        _pendingConfirmation = completion;
        Title = title;
        Message = message;
        Kind = ShellPromptKind.Confirmation;
        IsOpen = true;

        var frame = new DispatcherFrame();
        completion.Task.GetAwaiter().OnCompleted(() => frame.Continue = false);
        Dispatcher.PushFrame(frame);

        return completion.Task.GetAwaiter().GetResult();
    }

    private void ShowMessage(string title, string message, ShellPromptKind kind)
    {
        RunOnUiThread(() =>
        {
            ResolvePendingConfirmation(false);
            Title = title;
            Message = message;
            Kind = kind;
            IsOpen = true;
        });
    }

    private void Confirm()
    {
        if (Kind != ShellPromptKind.Confirmation)
        {
            Dismiss();
            return;
        }

        ResolvePendingConfirmation(true);
    }

    private void Cancel()
    {
        ResolvePendingConfirmation(false);
    }

    private void Dismiss()
    {
        Title = string.Empty;
        Message = string.Empty;
        IsOpen = false;
    }

    private void ResolvePendingConfirmation(bool result)
    {
        var pending = _pendingConfirmation;
        _pendingConfirmation = null;
        Dismiss();
        pending?.TrySetResult(result);
    }
}

