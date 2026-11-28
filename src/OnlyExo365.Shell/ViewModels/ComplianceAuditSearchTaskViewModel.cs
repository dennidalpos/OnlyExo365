using System.Globalization;
using System.Windows;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Shell.Localization;

namespace OnlyExo365.Shell.ViewModels;

public enum ComplianceAuditSearchTaskState
{
    Queued,
    Running,
    Completed,
    Failed
}

public sealed class ComplianceAuditSearchTaskViewModel : ViewModelBase
{
    private ComplianceAuditSearchTaskState _state = ComplianceAuditSearchTaskState.Queued;
    private int _progressPercent;
    private string? _workerStatusMessage;
    private int _resultCount;
    private string? _warning;
    private string? _errorMessage;
    private string? _correlationId;
    private IReadOnlyList<UnifiedAuditLogRecordDto> _results = Array.Empty<UnifiedAuditLogRecordDto>();

    public ComplianceAuditSearchTaskViewModel(int sequence, SearchUnifiedAuditLogRequest request)
    {
        Sequence = sequence;
        Request = request;
        RequestedAtLocal = DateTime.Now;
        WeakEventManager<LocalizationService, EventArgs>.AddHandler(
            LocalizationService.Instance,
            nameof(LocalizationService.CultureChanged),
            OnCultureChanged);
    }

    public int Sequence { get; }
    public string Title => Loc.GetFormat("Compliance.AuditTaskTitle", Sequence);
    public DateTime RequestedAtLocal { get; }
    public SearchUnifiedAuditLogRequest Request { get; }
    public string FilterSummary => BuildFilterSummary(Request);

    public ComplianceAuditSearchTaskState State
    {
        get => _state;
        private set
        {
            if (SetProperty(ref _state, value))
            {
                OnPropertyChanged(nameof(StateLabel));
                OnPropertyChanged(nameof(IsQueued));
                OnPropertyChanged(nameof(IsRunning));
                OnPropertyChanged(nameof(IsCompleted));
                OnPropertyChanged(nameof(IsFailed));
                OnPropertyChanged(nameof(StatusMessage));
                OnPropertyChanged(nameof(TaskSummary));
            }
        }
    }

    public string StateLabel => State switch
    {
        ComplianceAuditSearchTaskState.Queued => Loc.Get("Compliance.AuditStateQueued"),
        ComplianceAuditSearchTaskState.Running => Loc.Get("Compliance.AuditStateRunning"),
        ComplianceAuditSearchTaskState.Completed => Loc.Get("Compliance.AuditStateCompleted"),
        ComplianceAuditSearchTaskState.Failed => Loc.Get("Compliance.AuditStateFailed"),
        _ => Loc.Get("Compliance.AuditStateUnknown")
    };

    public bool IsQueued => State == ComplianceAuditSearchTaskState.Queued;
    public bool IsRunning => State == ComplianceAuditSearchTaskState.Running;
    public bool IsCompleted => State == ComplianceAuditSearchTaskState.Completed;
    public bool IsFailed => State == ComplianceAuditSearchTaskState.Failed;

    public int ProgressPercent
    {
        get => _progressPercent;
        private set
        {
            if (SetProperty(ref _progressPercent, Math.Clamp(value, 0, 100)))
            {
                OnPropertyChanged(nameof(ProgressText));
                OnPropertyChanged(nameof(TaskSummary));
            }
        }
    }

    public string ProgressText => FormatProgressPercent(ProgressPercent);

    public string StatusMessage => State switch
    {
        ComplianceAuditSearchTaskState.Running => _workerStatusMessage ?? Loc.Get("Presentation.RunningSearchUnifiedAuditLogInProgress"),
        _ => StateLabel
    };

    public int ResultCount
    {
        get => _resultCount;
        private set
        {
            if (SetProperty(ref _resultCount, value))
            {
                OnPropertyChanged(nameof(TaskSummary));
            }
        }
    }

    public string? Warning
    {
        get => _warning;
        private set
        {
            if (SetProperty(ref _warning, value))
            {
                OnPropertyChanged(nameof(HasWarning));
            }
        }
    }

    public bool HasWarning => !string.IsNullOrWhiteSpace(Warning);

    public string? ErrorMessage
    {
        get => _errorMessage;
        private set
        {
            if (SetProperty(ref _errorMessage, value))
            {
                OnPropertyChanged(nameof(HasError));
            }
        }
    }

    public bool HasError => !string.IsNullOrWhiteSpace(ErrorMessage);

    public string? CorrelationId
    {
        get => _correlationId;
        set => SetProperty(ref _correlationId, value);
    }

    public IReadOnlyList<UnifiedAuditLogRecordDto> Results
    {
        get => _results;
        private set => SetProperty(ref _results, value);
    }

    public string RequestedAtDisplay => RequestedAtLocal.ToString(
        "g", CultureInfo.GetCultureInfo(LocalizationService.Instance.CurrentLocale));

    public string TaskSummary => Loc.GetFormat("Compliance.AuditTaskSummary", Title, StateLabel, ProgressText, ResultCount);

    public void MarkRunning()
    {
        State = ComplianceAuditSearchTaskState.Running;
        ProgressPercent = Math.Max(ProgressPercent, 1);
        _workerStatusMessage = null;
        OnPropertyChanged(nameof(StatusMessage));
    }

    public void ApplyProgress(ProgressEventPayload payload)
    {
        State = ComplianceAuditSearchTaskState.Running;
        ProgressPercent = payload.PercentComplete;
        _workerStatusMessage = string.IsNullOrWhiteSpace(payload.StatusMessage) ? null : payload.StatusMessage;
        OnPropertyChanged(nameof(StatusMessage));
    }

    public void MarkCompleted(SearchUnifiedAuditLogResponse response, string? correlationId)
    {
        CorrelationId = correlationId;
        Results = response.Results.ToList();
        ResultCount = response.TotalCount > 0 ? response.TotalCount : response.Results.Count;
        Warning = response.Warning;
        ErrorMessage = null;
        ProgressPercent = 100;
        _workerStatusMessage = null;
        State = ComplianceAuditSearchTaskState.Completed;
    }

    public void MarkFailed(string message, string? correlationId)
    {
        CorrelationId = correlationId;
        Results = Array.Empty<UnifiedAuditLogRecordDto>();
        ResultCount = 0;
        Warning = null;
        ErrorMessage = message;
        ProgressPercent = 100;
        _workerStatusMessage = null;
        State = ComplianceAuditSearchTaskState.Failed;
    }

    private static string BuildFilterSummary(SearchUnifiedAuditLogRequest request)
    {
        var parts = new List<string>
        {
            $"{request.StartDate:yyyy-MM-dd} -> {request.EndDate:yyyy-MM-dd}",
            Loc.GetFormat("Compliance.AuditFilterMax", Math.Max(1, request.MaxResults))
        };

        if (request.UserIds.Count > 0)
        {
            parts.Add(Loc.GetFormat("Compliance.AuditFilterUsers", string.Join(", ", request.UserIds)));
        }

        if (request.Operations.Count > 0)
        {
            parts.Add(Loc.GetFormat("Compliance.AuditFilterOperations", string.Join(", ", request.Operations)));
        }

        if (request.ObjectIds.Count > 0)
        {
            parts.Add(Loc.GetFormat("Compliance.AuditFilterObjects", string.Join(", ", request.ObjectIds)));
        }

        if (!string.IsNullOrWhiteSpace(request.FreeText))
        {
            parts.Add(Loc.GetFormat("Compliance.AuditFilterText", request.FreeText.Trim()));
        }

        return string.Join(" | ", parts);
    }

    private void OnCultureChanged(object? sender, EventArgs e)
    {
        OnPropertyChanged(nameof(Title));
        OnPropertyChanged(nameof(FilterSummary));
        OnPropertyChanged(nameof(StateLabel));
        OnPropertyChanged(nameof(StatusMessage));
        OnPropertyChanged(nameof(RequestedAtDisplay));
        OnPropertyChanged(nameof(TaskSummary));
    }
}

