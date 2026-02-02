using System.Collections.ObjectModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

public class MessageTraceViewModel : ViewModelBase
{
    private readonly IWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;

    private CancellationTokenSource? _searchCts;

    private bool _isLoading;
    private string? _errorMessage;
    private string? _senderAddress;
    private string? _recipientAddress;
    private DateTime _startDate = DateTime.Today.AddDays(-7);
    private DateTime _endDate = DateTime.Today;
    private int _currentPage = 1;
    private int _pageSize = 100;
    private bool _hasMore;
    private int _totalCount;
    private double _loadingProgress;
    private string? _loadingStatus;

    public MessageTraceViewModel(IWorkerService workerService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;

        SearchCommand = new AsyncRelayCommand(SearchAsync, () => CanSearch);
        NextPageCommand = new AsyncRelayCommand(NextPageAsync, () => HasMore && !IsLoading);
        PreviousPageCommand = new AsyncRelayCommand(PreviousPageAsync, () => CurrentPage > 1 && !IsLoading);
    }

    public bool IsLoading
    {
        get => _isLoading;
        private set
        {
            if (SetProperty(ref _isLoading, value))
            {
                OnPropertyChanged(nameof(CanSearch));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

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

    public bool HasError => !string.IsNullOrEmpty(ErrorMessage);

    public string? SenderAddress
    {
        get => _senderAddress;
        set => SetProperty(ref _senderAddress, value);
    }

    public string? RecipientAddress
    {
        get => _recipientAddress;
        set => SetProperty(ref _recipientAddress, value);
    }

    public DateTime StartDate
    {
        get => _startDate;
        set => SetProperty(ref _startDate, value);
    }

    public DateTime EndDate
    {
        get => _endDate;
        set => SetProperty(ref _endDate, value);
    }

    public int CurrentPage
    {
        get => _currentPage;
        private set => SetProperty(ref _currentPage, value);
    }

    public int PageSize
    {
        get => _pageSize;
        set => SetProperty(ref _pageSize, value);
    }

    public bool HasMore
    {
        get => _hasMore;
        private set
        {
            if (SetProperty(ref _hasMore, value))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public int TotalCount
    {
        get => _totalCount;
        private set => SetProperty(ref _totalCount, value);
    }

    public double LoadingProgress
    {
        get => _loadingProgress;
        private set => SetProperty(ref _loadingProgress, value);
    }

    public string? LoadingStatus
    {
        get => _loadingStatus;
        private set => SetProperty(ref _loadingStatus, value);
    }

    public bool CanSearch => !IsLoading && _shellViewModel.IsExchangeConnected;

    public ObservableCollection<MessageTraceItemDto> Messages { get; } = new();

    public ICommand SearchCommand { get; }
    public ICommand NextPageCommand { get; }
    public ICommand PreviousPageCommand { get; }

    public async Task LoadAsync()
    {
        // Ensure commands re-evaluate CanExecute based on current connection state
        OnPropertyChanged(nameof(CanSearch));
        CommandManager.InvalidateRequerySuggested();

        if (!_shellViewModel.IsExchangeConnected)
        {
            ErrorMessage = "Non connesso a Exchange Online";
            return;
        }

        ErrorMessage = null;
    }

    private async Task SearchAsync(CancellationToken cancellationToken)
    {
        CurrentPage = 1;
        await FetchMessagesAsync(cancellationToken);
    }

    private async Task NextPageAsync(CancellationToken cancellationToken)
    {
        CurrentPage++;
        await FetchMessagesAsync(cancellationToken);
    }

    private async Task PreviousPageAsync(CancellationToken cancellationToken)
    {
        if (CurrentPage > 1)
        {
            CurrentPage--;
            await FetchMessagesAsync(cancellationToken);
        }
    }

    private async Task FetchMessagesAsync(CancellationToken cancellationToken)
    {
        _searchCts?.Cancel();
        _searchCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

        IsLoading = true;
        ErrorMessage = null;
        LoadingProgress = 0;
        LoadingStatus = "Ricerca messaggi...";

        try
        {
            var request = new GetMessageTraceRequest
            {
                SenderAddress = string.IsNullOrWhiteSpace(SenderAddress) ? null : SenderAddress.Trim(),
                RecipientAddress = string.IsNullOrWhiteSpace(RecipientAddress) ? null : RecipientAddress.Trim(),
                StartDate = StartDate,
                EndDate = EndDate.Date.AddDays(1).AddSeconds(-1),
                PageSize = PageSize,
                Page = CurrentPage
            };

            LoadingProgress = 30;
            LoadingStatus = "Interrogazione Exchange...";

            var result = await _workerService.GetMessageTraceAsync(
                request,
                eventHandler: evt =>
                {
                    if (evt.EventType == EventType.Progress)
                    {
                        var progress = JsonMessageSerializer.ExtractPayload<ProgressEventPayload>(evt.Payload);
                        if (progress != null)
                        {
                            RunOnUiThread(() =>
                            {
                                LoadingProgress = progress.PercentComplete;
                                LoadingStatus = progress.StatusMessage;
                            });
                        }
                    }
                },
                cancellationToken: _searchCts.Token);

            LoadingProgress = 90;
            LoadingStatus = "Elaborazione risultati...";

            if (result.IsSuccess && result.Value != null)
            {
                RunOnUiThread(() =>
                {
                    Messages.Clear();
                    foreach (var msg in result.Value.Messages)
                    {
                        Messages.Add(msg);
                    }
                    TotalCount = result.Value.TotalCount;
                    HasMore = result.Value.HasMore;
                });
                _shellViewModel.AddLog(LogLevel.Information, $"Message trace: {result.Value.Messages.Count} risultati trovati");
            }
            else if (result.WasCancelled)
            {
            }
            else
            {
                ErrorMessage = result.Error?.Message ?? "Impossibile recuperare la traccia messaggi";
                _shellViewModel.AddLog(LogLevel.Error, $"Message trace failed: {ErrorMessage}");
            }

            LoadingProgress = 100;
            LoadingStatus = null;
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Message trace error: {ex.Message}");
        }
        finally
        {
            IsLoading = false;
            LoadingProgress = 0;
            LoadingStatus = null;
        }
    }

    public void Cancel()
    {
        _searchCts?.Cancel();
    }
}
