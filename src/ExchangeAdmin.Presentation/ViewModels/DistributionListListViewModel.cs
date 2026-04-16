using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;
using System.Windows.Threading;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Paging;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;

namespace ExchangeAdmin.Presentation.ViewModels;

public sealed class DistributionListListViewModel : ViewModelBase
{
    private const int PageSize = PagingDefaults.DefaultPageSize;

    private readonly IDistributionListsWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;
    private readonly DispatcherTimer _searchDebounceTimer;
    private readonly Action<DistributionListItemDto?> _viewDetails;
    private readonly Action<string?> _setErrorMessage;

    private CancellationTokenSource? _loadCts;
    private bool _isLoading;
    private bool _hasMore;
    private bool _includeDynamic = true;
    private int _totalCount;
    private bool _isTotalCountExact = true;
    private int _currentSkip;
    private string? _searchQuery;
    private string _selectedGroupTypeFilter = "All";
    private DistributionListItemDto? _selectedItem;

    public DistributionListListViewModel(
        IDistributionListsWorkerService workerService,
        ShellViewModel shellViewModel,
        Action<DistributionListItemDto?> viewDetails,
        Action<string?> setErrorMessage)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;
        _viewDetails = viewDetails;
        _setErrorMessage = setErrorMessage;

        GroupTypeFilters = new ObservableCollection<string>(new[]
        {
            "All",
            "Distribution",
            "MailSecurity",
            "Microsoft365",
            "Dynamic"
        });

        _searchDebounceTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(300)
        };
        _searchDebounceTimer.Tick += OnSearchDebounceElapsed;

        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
        LoadMoreCommand = new AsyncRelayCommand(LoadMoreAsync, () => CanLoadMore);
        CancelCommand = new RelayCommand(Cancel, () => IsLoading);
        ViewDetailsCommand = new RelayCommand<DistributionListItemDto>(item => _viewDetails(item), item => item != null);
    }

    public ObservableCollection<string> GroupTypeFilters { get; }
    public ObservableCollection<DistributionListItemDto> DistributionLists { get; } = new();
    public CollectionLoadProgressViewModel LoadProgress { get; } = new();

    public bool IsLoading
    {
        get => _isLoading;
        private set
        {
            if (SetProperty(ref _isLoading, value))
            {
                OnPropertyChanged(nameof(CanRefresh));
                OnPropertyChanged(nameof(CanLoadMore));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string? SearchQuery
    {
        get => _searchQuery;
        set
        {
            if (SetProperty(ref _searchQuery, value))
            {
                _searchDebounceTimer.Stop();
                _searchDebounceTimer.Start();
            }
        }
    }

    public string SelectedGroupTypeFilter
    {
        get => _selectedGroupTypeFilter;
        set
        {
            var normalized = value;
            if (string.Equals(normalized, "Dynamic", StringComparison.OrdinalIgnoreCase) && !CanIncludeDynamicFilter)
            {
                normalized = "All";
            }

            if (SetProperty(ref _selectedGroupTypeFilter, normalized))
            {
                TriggerRefreshFromUi();
            }
        }
    }

    public bool IncludeDynamic
    {
        get => _includeDynamic;
        set
        {
            var normalized = value && CanIncludeDynamicFilter;
            if (SetProperty(ref _includeDynamic, normalized))
            {
                if (!normalized && string.Equals(SelectedGroupTypeFilter, "Dynamic", StringComparison.OrdinalIgnoreCase))
                {
                    SelectedGroupTypeFilter = "All";
                }
                else
                {
                    TriggerRefreshFromUi();
                }
            }
        }
    }

    public int TotalCount
    {
        get => _totalCount;
        private set => SetProperty(ref _totalCount, value);
    }

    public bool IsTotalCountExact
    {
        get => _isTotalCountExact;
        private set
        {
            if (SetProperty(ref _isTotalCountExact, value))
            {
                OnPropertyChanged(nameof(StatusText));
            }
        }
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

    public DistributionListItemDto? SelectedItem
    {
        get => _selectedItem;
        set
        {
            if (SetProperty(ref _selectedItem, value) && value != null)
            {
                _viewDetails(value);
            }
        }
    }

    public bool CanRefresh => !IsLoading && _shellViewModel.IsExchangeConnected;
    public bool CanLoadMore => !IsLoading && HasMore && _shellViewModel.IsExchangeConnected;
    public bool CanIncludeDynamicFilter => _shellViewModel.IsFeatureAvailable(f => f.CanGetDynamicDistributionGroup);
    public string StatusText => IsTotalCountExact
        ? $"{DistributionLists.Count} of {TotalCount} groups"
        : $"{DistributionLists.Count}+ groups loaded";

    public ICommand RefreshCommand { get; }
    public ICommand LoadMoreCommand { get; }
    public ICommand CancelCommand { get; }
    public ICommand ViewDetailsCommand { get; }

    public async Task LoadAsync(CancellationToken cancellationToken)
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            DistributionLists.Clear();
            _setErrorMessage("Not connected to Exchange Online");
            return;
        }

        if (!CanIncludeDynamicFilter && IncludeDynamic)
        {
            IncludeDynamic = false;
        }

        await RefreshAsync(cancellationToken);
    }

    public void Cancel() => _loadCts?.Cancel();

    public void HandleShellPropertyChanged(PropertyChangedEventArgs e)
    {
        if (e.PropertyName is not (nameof(ShellViewModel.Capabilities) or nameof(ShellViewModel.ExchangeState) or nameof(ShellViewModel.IsExchangeConnected)))
        {
            return;
        }

        OnPropertyChanged(nameof(CanRefresh));
        OnPropertyChanged(nameof(CanLoadMore));
        OnPropertyChanged(nameof(CanIncludeDynamicFilter));

        if (!CanIncludeDynamicFilter)
        {
            if (IncludeDynamic)
            {
                IncludeDynamic = false;
            }

            if (string.Equals(SelectedGroupTypeFilter, "Dynamic", StringComparison.OrdinalIgnoreCase))
            {
                SelectedGroupTypeFilter = "All";
            }
        }

        CommandManager.InvalidateRequerySuggested();
    }

    private void OnSearchDebounceElapsed(object? sender, EventArgs e)
    {
        _searchDebounceTimer.Stop();
        TriggerRefreshFromUi();
    }

    private void TriggerRefreshFromUi() => _ = SafeRefreshAsync();

    private async Task SafeRefreshAsync()
    {
        try
        {
            await RefreshAsync(CancellationToken.None);
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Group refresh failed: {ex.Message}");
        }
    }

    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        _loadCts?.Cancel();
        _loadCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

        IsLoading = true;
        LoadProgress.Start("Loading groups...", "groups");
        _setErrorMessage(null);
        var refreshPageSize = GetRefreshPageSize(DistributionLists.Count);
        _currentSkip = 0;

        try
        {
            var result = await _workerService.GetDistributionListsAsync(
                BuildListRequest(0, refreshPageSize),
                eventHandler: HandleLoadProgressEvent,
                cancellationToken: _loadCts.Token);
            if (result.IsSuccess && result.Value != null)
            {
                DistributionLists.ReplaceAll(result.Value.DistributionLists);
                TotalCount = result.Value.TotalCount;
                IsTotalCountExact = result.Value.IsTotalCountExact;
                HasMore = result.Value.HasMore;
                _currentSkip = DistributionLists.Count;
                OnPropertyChanged(nameof(StatusText));
                return;
            }

            if (!result.WasCancelled)
            {
                var errorMessage = result.Error?.Message ?? "Unable to load distribution lists.";
                _setErrorMessage(errorMessage);
                _shellViewModel.AddLog(LogLevel.Error, $"Group load failed: {errorMessage}");
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            _setErrorMessage(ex.Message);
            _shellViewModel.AddLog(LogLevel.Error, $"Group load error: {ex.Message}");
        }
        finally
        {
            LoadProgress.Reset();
            IsLoading = false;
        }
    }

    private async Task LoadMoreAsync(CancellationToken cancellationToken)
    {
        if (!CanLoadMore)
        {
            return;
        }

        _loadCts?.Cancel();
        _loadCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        IsLoading = true;
        LoadProgress.Start("Loading groups...", "groups");
        _setErrorMessage(null);

        try
        {
            var result = await _workerService.GetDistributionListsAsync(
                BuildListRequest(_currentSkip),
                eventHandler: HandleLoadProgressEvent,
                cancellationToken: _loadCts.Token);
            if (result.IsSuccess && result.Value != null)
            {
                foreach (var item in result.Value.DistributionLists)
                {
                    DistributionLists.Add(item);
                }

                TotalCount = result.Value.TotalCount;
                IsTotalCountExact = result.Value.IsTotalCountExact;
                HasMore = result.Value.HasMore;
                _currentSkip = DistributionLists.Count;
                OnPropertyChanged(nameof(StatusText));
                return;
            }

            if (!result.WasCancelled)
            {
                _setErrorMessage(result.Error?.Message ?? "Unable to load more distribution lists.");
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            _setErrorMessage(ex.Message);
        }
        finally
        {
            LoadProgress.Reset();
            IsLoading = false;
        }
    }

    private GetDistributionListsRequest BuildListRequest(int skip, int? pageSize = null)
    {
        return new GetDistributionListsRequest
        {
            SearchQuery = string.IsNullOrWhiteSpace(SearchQuery) ? null : SearchQuery.Trim(),
            GroupTypeFilter = string.Equals(SelectedGroupTypeFilter, "All", StringComparison.OrdinalIgnoreCase) ? null : SelectedGroupTypeFilter,
            IncludeDynamic = IncludeDynamic && CanIncludeDynamicFilter,
            PageSize = pageSize ?? PageSize,
            Skip = skip
        };
    }

    private static int GetRefreshPageSize(int loadedCount)
        => Math.Max(PageSize, loadedCount);

    private void HandleLoadProgressEvent(EventEnvelope evt)
    {
        _ = RunOnUiThreadAsync(() => LoadProgress.Apply(evt));
    }
}

