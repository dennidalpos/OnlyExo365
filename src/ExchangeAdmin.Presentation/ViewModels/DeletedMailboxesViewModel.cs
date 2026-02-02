using System.Collections.ObjectModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

public class DeletedMailboxesViewModel : ViewModelBase
{
    private const int PageSize = 50;

    private readonly IWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;

    private CancellationTokenSource? _loadCts;
    private bool _isLoading;
    private string? _errorMessage;
    private string? _upnQuery;
    private string? _activeSearchQuery;
    private int _totalCount;
    private int _currentSkip;
    private bool _hasMore;

    public DeletedMailboxesViewModel(IWorkerService workerService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;

        LoadAllCommand = new AsyncRelayCommand(LoadAllAsync, () => CanRefresh);
        CheckUpnCommand = new AsyncRelayCommand(CheckUpnAsync, () => CanCheckUpn);
        LoadMoreCommand = new AsyncRelayCommand(LoadMoreAsync, () => CanLoadMore);
        CancelCommand = new RelayCommand(Cancel, () => IsLoading);
    }

    public ObservableCollection<DeletedMailboxItemDto> Mailboxes { get; } = new();

    public bool IsLoading
    {
        get => _isLoading;
        private set
        {
            if (SetProperty(ref _isLoading, value))
            {
                OnPropertyChanged(nameof(CanRefresh));
                OnPropertyChanged(nameof(CanCheckUpn));
                OnPropertyChanged(nameof(CanLoadMore));
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

    public string? UpnQuery
    {
        get => _upnQuery;
        set
        {
            if (SetProperty(ref _upnQuery, value))
            {
                OnPropertyChanged(nameof(CanCheckUpn));
            }
        }
    }

    public int TotalCount
    {
        get => _totalCount;
        private set => SetProperty(ref _totalCount, value);
    }

    public bool HasMore
    {
        get => _hasMore;
        private set
        {
            if (SetProperty(ref _hasMore, value))
            {
                OnPropertyChanged(nameof(CanLoadMore));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string StatusText => $"{Mailboxes.Count} di {TotalCount} deleted mailbox";

    public bool CanRefresh => !IsLoading && _shellViewModel.IsExchangeConnected;
    public bool CanCheckUpn => !IsLoading && _shellViewModel.IsExchangeConnected && !string.IsNullOrWhiteSpace(UpnQuery);
    public bool CanLoadMore => !IsLoading && HasMore && _shellViewModel.IsExchangeConnected;

    public ICommand LoadAllCommand { get; }
    public ICommand CheckUpnCommand { get; }
    public ICommand LoadMoreCommand { get; }
    public ICommand CancelCommand { get; }

    public async Task LoadAsync()
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            await RunOnUiThreadAsync(() => Mailboxes.Clear());
            ErrorMessage = "Non connesso a Exchange Online";
            return;
        }

        await RefreshAsync(CancellationToken.None);
    }

    private async Task LoadAllAsync(CancellationToken cancellationToken)
    {
        _activeSearchQuery = null;
        await RefreshAsync(cancellationToken);
    }

    private async Task CheckUpnAsync(CancellationToken cancellationToken)
    {
        _activeSearchQuery = UpnQuery?.Trim();
        await RefreshAsync(cancellationToken);
    }

    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        _loadCts?.Cancel();
        if (!_shellViewModel.IsExchangeConnected)
        {
            IsLoading = false;
            await RunOnUiThreadAsync(() => Mailboxes.Clear());
            ErrorMessage = "Non connesso a Exchange Online";
            TotalCount = 0;
            HasMore = false;
            _currentSkip = 0;
            OnPropertyChanged(nameof(StatusText));
            return;
        }

        _loadCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

        IsLoading = true;
        ErrorMessage = null;
        _currentSkip = 0;

        try
        {
            var request = new GetDeletedMailboxesRequest
            {
                SearchQuery = _activeSearchQuery,
                IncludeInactive = true,
                IncludeSoftDeleted = true,
                PageSize = PageSize,
                Skip = 0
            };

            var result = await _workerService.GetDeletedMailboxesAsync(
                request,
                eventHandler: null,
                cancellationToken: _loadCts.Token);

            if (result.IsSuccess && result.Value != null)
            {
                await RunOnUiThreadAsync(() =>
                {
                    Mailboxes.ReplaceAll(result.Value.Mailboxes);
                    TotalCount = result.Value.TotalCount;
                    HasMore = result.Value.HasMore;
                    _currentSkip = result.Value.Mailboxes.Count;
                    OnPropertyChanged(nameof(StatusText));
                });
            }
            else if (!result.WasCancelled)
            {
                var errorDetails = result.Error != null
                    ? $"{result.Error.Code}: {result.Error.Message}"
                    : "Failed to load deleted mailboxes (no error details)";
                ErrorMessage = errorDetails;
                _shellViewModel.AddLog(LogLevel.Error, $"Deleted mailbox load failed: {errorDetails}");
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            ErrorMessage = $"Exception: {ex.GetType().Name} - {ex.Message}";
            _shellViewModel.AddLog(LogLevel.Error, $"Deleted mailbox exception: {ex.GetType().Name} - {ex.Message}");
        }
        finally
        {
            IsLoading = false;
        }
    }

    private async Task LoadMoreAsync(CancellationToken cancellationToken)
    {
        if (!HasMore || IsLoading) return;

        _loadCts?.Cancel();
        if (!_shellViewModel.IsExchangeConnected)
        {
            IsLoading = false;
            ErrorMessage = "Non connesso a Exchange Online";
            return;
        }

        _loadCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

        IsLoading = true;
        ErrorMessage = null;

        try
        {
            var request = new GetDeletedMailboxesRequest
            {
                SearchQuery = _activeSearchQuery,
                IncludeInactive = true,
                IncludeSoftDeleted = true,
                PageSize = PageSize,
                Skip = _currentSkip
            };

            var result = await _workerService.GetDeletedMailboxesAsync(
                request,
                eventHandler: null,
                cancellationToken: _loadCts.Token);

            if (result.IsSuccess && result.Value != null)
            {
                await RunOnUiThreadAsync(() =>
                {
                    foreach (var mailbox in result.Value.Mailboxes)
                    {
                        Mailboxes.Add(mailbox);
                    }

                    HasMore = result.Value.HasMore;
                    _currentSkip += result.Value.Mailboxes.Count;
                    OnPropertyChanged(nameof(StatusText));
                });
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? "Failed to load more deleted mailboxes";
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsLoading = false;
        }
    }

    public void Cancel()
    {
        _loadCts?.Cancel();
    }
}
