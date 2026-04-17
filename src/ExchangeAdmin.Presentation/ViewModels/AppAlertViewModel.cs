namespace ExchangeAdmin.Presentation.ViewModels;

public enum AppAlertSeverity
{
    Error,
    Warning,
    Info
}

public sealed class AppAlertViewModel : ViewModelBase
{
    private AppAlertSeverity _severity = AppAlertSeverity.Info;
    private string _title = string.Empty;
    private string _message = string.Empty;
    private string? _details;
    private string? _source;
    private int _priority;

    public AppAlertSeverity Severity
    {
        get => _severity;
        private set => SetProperty(ref _severity, value);
    }

    public string Title
    {
        get => _title;
        private set
        {
            if (SetProperty(ref _title, value))
            {
                OnPropertyChanged(nameof(IsVisible));
            }
        }
    }

    public string Message
    {
        get => _message;
        private set
        {
            if (SetProperty(ref _message, value))
            {
                OnPropertyChanged(nameof(IsVisible));
            }
        }
    }

    public string? Details
    {
        get => _details;
        private set
        {
            if (SetProperty(ref _details, value))
            {
                OnPropertyChanged(nameof(HasDetails));
            }
        }
    }

    public string? Source
    {
        get => _source;
        private set
        {
            if (SetProperty(ref _source, value))
            {
                OnPropertyChanged(nameof(HasSource));
            }
        }
    }

    public int Priority
    {
        get => _priority;
        private set => SetProperty(ref _priority, value);
    }

    public bool IsVisible => !string.IsNullOrWhiteSpace(Title) && !string.IsNullOrWhiteSpace(Message);
    public bool HasDetails => !string.IsNullOrWhiteSpace(Details);
    public bool HasSource => !string.IsNullOrWhiteSpace(Source);

    public void Update(
        AppAlertSeverity severity,
        string title,
        string message,
        string? details = null,
        string? source = null,
        int priority = 0)
    {
        Severity = severity;
        Title = title;
        Message = message;
        Details = string.IsNullOrWhiteSpace(details) ? null : details.Trim();
        Source = string.IsNullOrWhiteSpace(source) ? null : source.Trim();
        Priority = priority;
    }

    public void Clear()
    {
        Severity = AppAlertSeverity.Info;
        Title = string.Empty;
        Message = string.Empty;
        Details = null;
        Source = null;
        Priority = 0;
    }
}
