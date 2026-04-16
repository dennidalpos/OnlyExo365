using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Presentation.ViewModels;

public sealed class ShellProgressViewModel : ViewModelBase
{
    private readonly HashSet<string> _backgroundOperationCorrelationIds = new(StringComparer.Ordinal);
    private bool _isGlobalOperationRunning;
    private int _globalProgress;
    private string? _globalStatus;
    private int? _globalCurrentItem;
    private int? _globalTotalItems;

    public bool IsGlobalOperationRunning
    {
        get => _isGlobalOperationRunning;
        internal set => SetProperty(ref _isGlobalOperationRunning, value);
    }

    public int GlobalProgress
    {
        get => _globalProgress;
        set
        {
            if (SetProperty(ref _globalProgress, value))
            {
                OnPropertyChanged(nameof(GlobalProgressPercentText));
            }
        }
    }

    public string? GlobalStatus
    {
        get => _globalStatus;
        set => SetProperty(ref _globalStatus, value);
    }

    public int? GlobalCurrentItem
    {
        get => _globalCurrentItem;
        internal set
        {
            if (SetProperty(ref _globalCurrentItem, value))
            {
                OnPropertyChanged(nameof(HasGlobalItemProgress));
                OnPropertyChanged(nameof(GlobalProgressCountText));
            }
        }
    }

    public int? GlobalTotalItems
    {
        get => _globalTotalItems;
        internal set
        {
            if (SetProperty(ref _globalTotalItems, value))
            {
                OnPropertyChanged(nameof(HasGlobalItemProgress));
                OnPropertyChanged(nameof(GlobalProgressCountText));
            }
        }
    }

    public bool HasGlobalItemProgress => GlobalCurrentItem.HasValue;

    public string GlobalProgressPercentText => FormatProgressPercent(GlobalProgress);
    public string? GlobalProgressCountText => FormatProgressCount(GlobalCurrentItem, GlobalTotalItems, "items");

    public void RegisterBackgroundOperation(string correlationId)
    {
        if (!string.IsNullOrWhiteSpace(correlationId))
        {
            _backgroundOperationCorrelationIds.Add(correlationId);
        }
    }

    public void UnregisterBackgroundOperation(string correlationId)
    {
        if (!string.IsNullOrWhiteSpace(correlationId))
        {
            _backgroundOperationCorrelationIds.Remove(correlationId);
        }
    }

    public void Apply(string correlationId, ProgressEventPayload payload)
    {
        if (!string.IsNullOrWhiteSpace(correlationId) && _backgroundOperationCorrelationIds.Contains(correlationId))
        {
            return;
        }

        GlobalProgress = payload.PercentComplete;
        GlobalStatus = payload.StatusMessage;
        IsGlobalOperationRunning = payload.PercentComplete < 100;
        GlobalCurrentItem = payload.CurrentItem;
        GlobalTotalItems = payload.TotalItems;
    }

    public void Reset()
    {
        IsGlobalOperationRunning = false;
        GlobalProgress = 0;
        GlobalStatus = null;
        GlobalCurrentItem = null;
        GlobalTotalItems = null;
    }
}
