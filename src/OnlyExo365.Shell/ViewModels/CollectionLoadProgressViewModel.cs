using OnlyExo365.Contracts.Messages;

namespace OnlyExo365.Shell.ViewModels;

public sealed class CollectionLoadProgressViewModel : ViewModelBase
{
    private bool _isVisible;
    private double _progress;
    private string _status = string.Empty;
    private int? _currentItem;
    private int? _totalItems;
    private string _fallbackCompletedLabel = "items";

    public bool IsVisible
    {
        get => _isVisible;
        private set => SetProperty(ref _isVisible, value);
    }

    public double Progress
    {
        get => _progress;
        private set
        {
            if (SetProperty(ref _progress, value))
            {
                OnPropertyChanged(nameof(PercentText));
            }
        }
    }

    public string Status
    {
        get => _status;
        private set => SetProperty(ref _status, value);
    }

    public int? CurrentItem
    {
        get => _currentItem;
        private set
        {
            if (SetProperty(ref _currentItem, value))
            {
                OnPropertyChanged(nameof(HasCount));
                OnPropertyChanged(nameof(CountText));
            }
        }
    }

    public int? TotalItems
    {
        get => _totalItems;
        private set
        {
            if (SetProperty(ref _totalItems, value))
            {
                OnPropertyChanged(nameof(HasCount));
                OnPropertyChanged(nameof(CountText));
            }
        }
    }

    public bool HasCount => CurrentItem.HasValue;

    public string PercentText => FormatProgressPercent(Progress);

    public string? CountText => FormatProgressCount(CurrentItem, TotalItems, _fallbackCompletedLabel);

    public void Start(string status, string fallbackCompletedLabel = "items")
    {
        _fallbackCompletedLabel = fallbackCompletedLabel;
        IsVisible = true;
        Progress = 0;
        Status = status;
        CurrentItem = null;
        TotalItems = null;
    }

    public void Apply(EventEnvelope? evt)
    {
        if (evt?.EventType != EventType.Progress)
        {
            return;
        }

        var progress = JsonMessageSerializer.ExtractPayload<ProgressEventPayload>(evt.Payload);
        if (progress == null)
        {
            return;
        }

        Apply(progress);
    }

    public void Apply(ProgressEventPayload progress)
    {
        IsVisible = true;
        Progress = progress.PercentComplete;
        Status = string.IsNullOrWhiteSpace(progress.StatusMessage) ? Status : progress.StatusMessage!;
        CurrentItem = progress.CurrentItem;
        TotalItems = progress.TotalItems;
    }

    public void Complete(string status, int? currentItem = null, int? totalItems = null)
    {
        IsVisible = true;
        Progress = 100;
        Status = status;
        CurrentItem = currentItem;
        TotalItems = totalItems;
    }

    public void Reset()
    {
        IsVisible = false;
        Progress = 0;
        Status = string.Empty;
        CurrentItem = null;
        TotalItems = null;
    }
}

