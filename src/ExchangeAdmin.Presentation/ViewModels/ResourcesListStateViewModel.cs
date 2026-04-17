using System.Collections.ObjectModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Paging;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

public sealed class ResourcesListStateViewModel : ViewModelBase
{
    private const NavigationPage AlertPage = NavigationPage.Resources;
    private const int PageSize = PagingDefaults.DefaultPageSize;

    private readonly IResourcesWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;
    private readonly DebounceHelper _searchDebounce = new();
    private readonly Action<ResourceMailboxListItemDto?> _onSelectionChanged;
    private readonly Action<string?> _setErrorMessage;

    private bool _isLoading;
    private string? _searchQuery;
    private string _resourceTypeFilter = "All";
    private int _totalCount;
    private bool _hasMore;
    private int _currentSkip;
    private ResourceMailboxListItemDto? _selectedResource;

    public ResourcesListStateViewModel(
        IResourcesWorkerService workerService,
        ShellViewModel shellViewModel,
        Action<ResourceMailboxListItemDto?> onSelectionChanged,
        Action<string?> setErrorMessage)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;
        _onSelectionChanged = onSelectionChanged;
        _setErrorMessage = setErrorMessage;

        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
        LoadMoreCommand = new AsyncRelayCommand(LoadMoreAsync, () => CanLoadMore);
    }

    public ObservableCollection<ResourceMailboxListItemDto> Resources { get; } = new();
    public CollectionLoadProgressViewModel LoadProgress { get; } = new();

    public IReadOnlyList<string> ResourceTypeFilters { get; } = ["All", "Room", "Equipment"];

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
                _searchDebounce.Debounce(TriggerRefreshFromUi, 300);
            }
        }
    }

    public string ResourceTypeFilter
    {
        get => _resourceTypeFilter;
        set
        {
            if (SetProperty(ref _resourceTypeFilter, value))
            {
                TriggerRefreshFromUi();
            }
        }
    }

    public int TotalCount
    {
        get => _totalCount;
        private set
        {
            if (SetProperty(ref _totalCount, value))
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

    public string StatusText => $"{Resources.Count} of {TotalCount} resources";

    public ResourceMailboxListItemDto? SelectedResource
    {
        get => _selectedResource;
        set
        {
            if (SetProperty(ref _selectedResource, value) && value != null)
            {
                _onSelectionChanged(value);
            }
        }
    }

    public bool CanRefresh => !IsLoading && _shellViewModel.IsExchangeConnected;
    public bool CanLoadMore => !IsLoading && HasMore && _shellViewModel.IsExchangeConnected;

    public ICommand RefreshCommand { get; }
    public ICommand LoadMoreCommand { get; }

    public async Task LoadAsync(CancellationToken cancellationToken)
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            Resources.Clear();
            TotalCount = 0;
            HasMore = false;
            _setErrorMessage(null);
            _shellViewModel.ClearPageAlert(AlertPage);
            return;
        }

        if (Resources.Count == 0)
        {
            await RefreshAsync(cancellationToken);
        }
    }

    public void ClearSelection()
    {
        SelectedResource = null;
    }

    public async Task RefreshAsync(CancellationToken cancellationToken)
    {
        var hasExistingResources = Resources.Count > 0;

        if (!_shellViewModel.IsExchangeConnected)
        {
            Resources.Clear();
            TotalCount = 0;
            HasMore = false;
            _setErrorMessage(null);
            _shellViewModel.ClearPageAlert(AlertPage);
            return;
        }

        IsLoading = true;
        LoadProgress.Start("Loading resource mailboxes...", "resources");
        _setErrorMessage(null);
        _shellViewModel.ClearPageAlert(AlertPage);
        var refreshPageSize = GetRefreshPageSize(Resources.Count);
        _currentSkip = 0;

        try
        {
            var result = await _workerService.GetResourceMailboxesAsync(
                BuildRequest(0, refreshPageSize),
                eventHandler: HandleLoadProgressEvent,
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                var errorMessage = result.Error?.Message ?? "Unable to load resources";
                if (hasExistingResources)
                {
                    _setErrorMessage(errorMessage);
                }
                else
                {
                    _setErrorMessage(null);
                    _shellViewModel.ShowPageLoadFailedAlert(AlertPage, errorMessage);
                }
                return;
            }

            Resources.ReplaceAll(result.Value.Resources);
            TotalCount = result.Value.TotalCount;
            HasMore = result.Value.HasMore;
            _currentSkip = Resources.Count;
            _shellViewModel.ClearPageAlert(AlertPage);
        }
        catch (Exception ex)
        {
            if (hasExistingResources)
            {
                _setErrorMessage(ex.Message);
            }
            else
            {
                _setErrorMessage(null);
                _shellViewModel.ShowPageLoadFailedAlert(AlertPage, ex.Message);
            }
        }
        finally
        {
            LoadProgress.Reset();
            IsLoading = false;
        }
    }

    public async Task LoadMoreAsync(CancellationToken cancellationToken)
    {
        if (!HasMore)
        {
            return;
        }

        IsLoading = true;
        LoadProgress.Start("Loading resource mailboxes...", "resources");
        _setErrorMessage(null);

        try
        {
            var result = await _workerService.GetResourceMailboxesAsync(
                BuildRequest(_currentSkip),
                eventHandler: HandleLoadProgressEvent,
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                _setErrorMessage(result.Error?.Message ?? "Unable to load resources");
                return;
            }

            foreach (var item in result.Value.Resources)
            {
                Resources.Add(item);
            }

            TotalCount = result.Value.TotalCount;
            HasMore = result.Value.HasMore;
            _currentSkip = Resources.Count;
            OnPropertyChanged(nameof(StatusText));
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

    private void TriggerRefreshFromUi() => _ = SafeRefreshAsync();

    private async Task SafeRefreshAsync()
    {
        try
        {
            await RefreshAsync(CancellationToken.None);
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Resources refresh failed: {ex.Message}", "Resources");
        }
    }

    private GetResourceMailboxesRequest BuildRequest(int skip, int? pageSize = null)
    {
        return new GetResourceMailboxesRequest
        {
            ResourceType = NormalizeFilter(ResourceTypeFilter),
            SearchQuery = string.IsNullOrWhiteSpace(SearchQuery) ? null : SearchQuery.Trim(),
            PageSize = pageSize ?? PageSize,
            Skip = skip
        };
    }

    private static string? NormalizeFilter(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || string.Equals(value, "All", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return value;
    }

    private static int GetRefreshPageSize(int loadedCount)
        => Math.Max(PageSize, loadedCount);

    private void HandleLoadProgressEvent(EventEnvelope evt)
    {
        _ = RunOnUiThreadAsync(() => LoadProgress.Apply(evt));
    }
}

