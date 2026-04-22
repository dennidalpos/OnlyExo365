using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;

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
    private string _statusMessage = "Queued";
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
        Title = $"Audit #{sequence:000}";
        FilterSummary = BuildFilterSummary(request);
    }

    public int Sequence { get; }
    public string Title { get; }
    public DateTime RequestedAtLocal { get; }
    public SearchUnifiedAuditLogRequest Request { get; }
    public string FilterSummary { get; }

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
                OnPropertyChanged(nameof(TaskSummary));
            }
        }
    }

    public string StateLabel => State switch
    {
        ComplianceAuditSearchTaskState.Queued => "Queued",
        ComplianceAuditSearchTaskState.Running => "Running",
        ComplianceAuditSearchTaskState.Completed => "Completed",
        ComplianceAuditSearchTaskState.Failed => "Failed",
        _ => "Unknown"
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

    public string StatusMessage
    {
        get => _statusMessage;
        private set
        {
            if (SetProperty(ref _statusMessage, value))
            {
                OnPropertyChanged(nameof(TaskSummary));
            }
        }
    }

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

    public string RequestedAtDisplay => RequestedAtLocal.ToString("g");

    public string TaskSummary => $"{Title} | {StateLabel} | {ProgressText} | {ResultCount} rec";

    public void MarkRunning()
    {
        State = ComplianceAuditSearchTaskState.Running;
        ProgressPercent = Math.Max(ProgressPercent, 1);
        StatusMessage = "Running Search-UnifiedAuditLog...";
    }

    public void ApplyProgress(ProgressEventPayload payload)
    {
        State = ComplianceAuditSearchTaskState.Running;
        ProgressPercent = payload.PercentComplete;
        StatusMessage = string.IsNullOrWhiteSpace(payload.StatusMessage)
            ? "Running Search-UnifiedAuditLog..."
            : payload.StatusMessage!;
    }

    public void MarkCompleted(SearchUnifiedAuditLogResponse response, string? correlationId)
    {
        CorrelationId = correlationId;
        Results = response.Results.ToList();
        ResultCount = response.TotalCount > 0 ? response.TotalCount : response.Results.Count;
        Warning = response.Warning;
        ErrorMessage = null;
        ProgressPercent = 100;
        StatusMessage = "Completed";
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
        StatusMessage = "Failed";
        State = ComplianceAuditSearchTaskState.Failed;
    }

    private static string BuildFilterSummary(SearchUnifiedAuditLogRequest request)
    {
        var parts = new List<string>
        {
            $"{request.StartDate:yyyy-MM-dd} -> {request.EndDate:yyyy-MM-dd}",
            $"Max {Math.Max(1, request.MaxResults)}"
        };

        if (request.UserIds.Count > 0)
        {
            parts.Add($"Users: {string.Join(", ", request.UserIds)}");
        }

        if (request.Operations.Count > 0)
        {
            parts.Add($"Ops: {string.Join(", ", request.Operations)}");
        }

        if (request.ObjectIds.Count > 0)
        {
            parts.Add($"Objects: {string.Join(", ", request.ObjectIds)}");
        }

        if (!string.IsNullOrWhiteSpace(request.FreeText))
        {
            parts.Add($"Text: {request.FreeText.Trim()}");
        }

        return string.Join(" | ", parts);
    }
}

