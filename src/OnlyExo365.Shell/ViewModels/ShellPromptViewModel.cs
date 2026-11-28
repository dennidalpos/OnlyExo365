using System.Windows.Input;
using OnlyExo365.Shell.Services;
using OnlyExo365.Shell.Helpers;

namespace OnlyExo365.Shell.ViewModels;

public enum ShellPromptKind
{
    Information,
    Warning
}

public sealed class ShellPromptViewModel : ViewModelBase
{
    private bool _isOpen;
    private string _title = string.Empty;
    private string _message = string.Empty;
    private ShellPromptKind _kind = ShellPromptKind.Information;

    public ShellPromptViewModel()
    {
        CancelCommand = new RelayCommand(Dismiss, () => IsOpen);
    }

    public bool IsOpen
    {
        get => _isOpen;
        private set
        {
            if (SetProperty(ref _isOpen, value))
            {
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
                OnPropertyChanged(nameof(IsWarning));
                OnPropertyChanged(nameof(IsInformation));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool IsWarning => IsOpen && Kind == ShellPromptKind.Warning;
    public bool IsInformation => IsOpen && Kind == ShellPromptKind.Information;

    public ICommand CancelCommand { get; }

    public void ShowInformation(string title, string message)
        => ShowMessage(title, message, ShellPromptKind.Information);

    public void ShowWarning(string title, string message)
        => ShowMessage(title, message, ShellPromptKind.Warning);

    public bool ShowConfirmationBlocking(string title, string message)
        => ErrorDialogService.ShowConfirmation(title, message);

    private void ShowMessage(string title, string message, ShellPromptKind kind)
    {
        RunOnUiThread(() =>
        {
            Dismiss();
            Title = title;
            Message = message;
            Kind = kind;
            IsOpen = true;
        });
    }

    private void Dismiss()
    {
        Title = string.Empty;
        Message = string.Empty;
        IsOpen = false;
    }

}
