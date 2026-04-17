using System.Collections.ObjectModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Paging;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

public class DeletedMailboxesViewModel : ViewModelBase
{
    private const NavigationPage AlertPage = NavigationPage.DeletedMailboxes;
    private const int PageSize = PagingDefaults.DefaultPageSize;

    private readonly IWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;

    private CancellationTokenSource? _loadCts;
    private bool _isLoading;
    private string? _errorMessage;
    private string? _upnQuery;
    private string? _activeSearchQuery;
    private bool _includeSoftDeleted = true;
    private bool _includeInactive = true;
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
    public CollectionLoadProgressViewModel LoadProgress { get; } = new();

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


    public bool IncludeSoftDeleted
    {
        get => _includeSoftDeleted;
        set
        {
            if (SetProperty(ref _includeSoftDeleted, value))
            {
                _ = RefreshFromFiltersAsync();
            }
        }
    }

    public bool IncludeInactive
    {
        get => _includeInactive;
        set
        {
            if (SetProperty(ref _includeInactive, value))
            {
                _ = RefreshFromFiltersAsync();
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

    public string StatusText => $"{Mailboxes.Count} of {TotalCount} deleted mailboxes";

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
            ErrorMessage = null;
            _shellViewModel.ClearPageAlert(AlertPage);
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


    private async Task RefreshFromFiltersAsync()
    {
        try
        {
            await RefreshAsync(CancellationToken.None);
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Deleted mailbox refresh failed: {ex.Message}");
        }
    }

    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        _loadCts?.Cancel();
        var hasExistingMailboxes = Mailboxes.Count > 0;
        if (!_shellViewModel.IsExchangeConnected)
        {
            IsLoading = false;
            await RunOnUiThreadAsync(() => Mailboxes.Clear());
            ErrorMessage = null;
            TotalCount = 0;
            HasMore = false;
            _currentSkip = 0;
            _shellViewModel.ClearPageAlert(AlertPage);
            OnPropertyChanged(nameof(StatusText));
            return;
        }

        _loadCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

        IsLoading = true;
        LoadProgress.Start("Loading deleted mailboxes...", "mailbox");
        ErrorMessage = null;
        _shellViewModel.ClearPageAlert(AlertPage);
        var refreshPageSize = GetRefreshPageSize(Mailboxes.Count);
        _currentSkip = 0;

        try
        {
            var request = new GetDeletedMailboxesRequest
            {
                SearchQuery = _activeSearchQuery,
                IncludeInactive = _includeInactive,
                IncludeSoftDeleted = _includeSoftDeleted,
                PageSize = refreshPageSize,
                Skip = 0
            };

            var result = await _workerService.GetDeletedMailboxesAsync(
                request,
                eventHandler: HandleLoadProgressEvent,
                cancellationToken: _loadCts.Token);

            if (result.IsSuccess && result.Value != null)
            {
                await RunOnUiThreadAsync(() =>
                {
                    Mailboxes.ReplaceAll(result.Value.Mailboxes);
                    TotalCount = result.Value.TotalCount;
                    HasMore = result.Value.HasMore;
                    _currentSkip = Mailboxes.Count;
                    _shellViewModel.ClearPageAlert(AlertPage);
                    OnPropertyChanged(nameof(StatusText));
                });
            }
            else if (!result.WasCancelled)
            {
                var errorDetails = result.Error != null
                    ? $"{result.Error.Code}: {result.Error.Message}"
                    : "Failed to load deleted mailboxes (no error details)";
                ErrorMessage = hasExistingMailboxes ? errorDetails : null;
                if (!hasExistingMailboxes)
                {
                    _shellViewModel.ShowPageLoadFailedAlert(AlertPage, errorDetails);
                }
                _shellViewModel.AddLog(LogLevel.Error, $"Deleted mailbox load failed: {errorDetails}");
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            var errorDetails = $"Exception: {ex.GetType().Name} - {ex.Message}";
            ErrorMessage = hasExistingMailboxes ? errorDetails : null;
            if (!hasExistingMailboxes)
            {
                _shellViewModel.ShowPageLoadFailedAlert(AlertPage, errorDetails);
            }
            _shellViewModel.AddLog(LogLevel.Error, $"Deleted mailbox exception: {ex.GetType().Name} - {ex.Message}");
        }
        finally
        {
            LoadProgress.Reset();
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
            ErrorMessage = null;
            _shellViewModel.ClearPageAlert(AlertPage);
            return;
        }

        _loadCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

        IsLoading = true;
        LoadProgress.Start("Loading deleted mailboxes...", "mailbox");
        ErrorMessage = null;

        try
        {
            var request = new GetDeletedMailboxesRequest
            {
                SearchQuery = _activeSearchQuery,
                IncludeInactive = _includeInactive,
                IncludeSoftDeleted = _includeSoftDeleted,
                PageSize = PageSize,
                Skip = _currentSkip
            };

            var result = await _workerService.GetDeletedMailboxesAsync(
                request,
                eventHandler: HandleLoadProgressEvent,
                cancellationToken: _loadCts.Token);

            if (result.IsSuccess && result.Value != null)
            {
                await RunOnUiThreadAsync(() =>
                {
                    foreach (var mailbox in result.Value.Mailboxes)
                    {
                        Mailboxes.Add(mailbox);
                    }

                    TotalCount = result.Value.TotalCount;
                    HasMore = result.Value.HasMore;
                    _currentSkip = Mailboxes.Count;
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
            LoadProgress.Reset();
            IsLoading = false;
        }
    }

    public void Cancel()
    {
        _loadCts?.Cancel();
    }

    private void HandleLoadProgressEvent(EventEnvelope evt)
    {
        _ = RunOnUiThreadAsync(() => LoadProgress.Apply(evt));
    }

    private static int GetRefreshPageSize(int loadedCount)
        => Math.Max(PageSize, loadedCount);
}

