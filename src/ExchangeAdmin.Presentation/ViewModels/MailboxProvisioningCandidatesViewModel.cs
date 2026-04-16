using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Paging;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;

namespace ExchangeAdmin.Presentation.ViewModels;

public sealed class MailboxProvisioningCandidatesViewModel : ViewModelBase
{
    private const int PageSize = PagingDefaults.DefaultPageSize;

    private readonly IMailboxesWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;
    private readonly DebounceHelper _searchDebounce = new();

    private CancellationTokenSource? _loadCts;
    private bool _isLoading;
    private string? _errorMessage;
    private string? _searchQuery;
    private bool _onlyWithoutLicense = true;
    private bool _onlyWithoutMail = true;
    private int _totalCount;
    private int _currentSkip;
    private bool _hasMore;
    private MailboxProvisioningCandidateDto? _selectedCandidate;

    public MailboxProvisioningCandidatesViewModel(IMailboxesWorkerService workerService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;

        Licenses = new MailboxLicensesViewModel(
            workerService,
            shellViewModel,
            () => SelectedCandidate?.UserPrincipalName,
            () => SelectedCandidate?.DisplayName,
            RefreshAfterLicenseMutationAsync);

        Licenses.PropertyChanged += OnLicensesPropertyChanged;
        _shellViewModel.PropertyChanged += OnShellPropertyChanged;

        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
        LoadMoreCommand = new AsyncRelayCommand(LoadMoreAsync, () => CanLoadMore);
    }

    public ObservableCollection<MailboxProvisioningCandidateDto> Candidates { get; } = new();

    public MailboxLicensesViewModel Licenses { get; }
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

    public bool IsLicenseSaving => Licenses.IsLicenseSaving;

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

    public bool OnlyWithoutLicense
    {
        get => _onlyWithoutLicense;
        set
        {
            if (SetProperty(ref _onlyWithoutLicense, value))
            {
                TriggerRefreshFromUi();
            }
        }
    }

    public bool OnlyWithoutMail
    {
        get => _onlyWithoutMail;
        set
        {
            if (SetProperty(ref _onlyWithoutMail, value))
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
                OnPropertyChanged(nameof(StatusText));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public MailboxProvisioningCandidateDto? SelectedCandidate
    {
        get => _selectedCandidate;
        set
        {
            if (SetProperty(ref _selectedCandidate, value))
            {
                OnPropertyChanged(nameof(HasSelectedCandidate));
                OnPropertyChanged(nameof(SelectedCandidateSummary));
                _ = LoadSelectedCandidateLicensesAsync();
            }
        }
    }

    public bool HasSelectedCandidate => SelectedCandidate != null;

    public string StatusText => $"{Candidates.Count} of {TotalCount} users";

    public string SelectedCandidateSummary => SelectedCandidate == null
        ? "Select a member user to review and assign licenses."
        : $"{SelectedCandidate.DisplayName} ({SelectedCandidate.UserPrincipalName})";

    public bool CanRefresh => !IsLoading && _shellViewModel.IsExchangeConnected;

    public bool CanLoadMore => !IsLoading && HasMore && _shellViewModel.IsExchangeConnected;

    public ICommand RefreshCommand { get; }

    public ICommand LoadMoreCommand { get; }

    public async Task LoadAsync(CancellationToken cancellationToken = default)
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            Reset();
            ErrorMessage = "Not connected to Exchange Online";
            return;
        }

        await RefreshAsync(cancellationToken);
    }

    public void Reset()
    {
        _loadCts?.Cancel();
        Candidates.Clear();
        SelectedCandidate = null;
        Licenses.Reset();
        LoadProgress.Reset();
        ErrorMessage = null;
        TotalCount = 0;
        HasMore = false;
        _currentSkip = 0;
    }

    private void TriggerRefreshFromUi()
    {
        _ = SafeRefreshAsync();
    }

    private async Task SafeRefreshAsync()
    {
        try
        {
            await RefreshAsync(CancellationToken.None);
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Mailbox provisioning refresh failed: {ex.Message}");
        }
    }

    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        _loadCts?.Cancel();

        if (!_shellViewModel.IsExchangeConnected)
        {
            Reset();
            ErrorMessage = "Not connected to Exchange Online";
            return;
        }

        var selectedUpn = SelectedCandidate?.UserPrincipalName;
        _loadCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        IsLoading = true;
        LoadProgress.Start("Loading users Member...", "users");
        ErrorMessage = null;
        var refreshPageSize = GetRefreshPageSize(Candidates.Count);
        _currentSkip = 0;

        try
        {
            var result = await _workerService.GetMailboxProvisioningCandidatesAsync(
                BuildRequest(skip: 0, pageSize: refreshPageSize),
                eventHandler: HandleLoadProgressEvent,
                cancellationToken: _loadCts.Token);

            if (result.IsSuccess && result.Value != null)
            {
                await RunOnUiThreadAsync(() =>
                {
                    Candidates.ReplaceAll(result.Value.Candidates);
                    TotalCount = result.Value.TotalCount;
                    HasMore = result.Value.HasMore;
                    _currentSkip = Candidates.Count;
                    SelectedCandidate = ResolveSelection(selectedUpn);
                });
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? "Failed to load member users";
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

    private async Task LoadMoreAsync(CancellationToken cancellationToken)
    {
        if (!HasMore || IsLoading)
        {
            return;
        }

        _loadCts?.Cancel();
        _loadCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        IsLoading = true;
        LoadProgress.Start("Loading users Member...", "users");
        ErrorMessage = null;

        try
        {
            var result = await _workerService.GetMailboxProvisioningCandidatesAsync(
                BuildRequest(_currentSkip),
                eventHandler: HandleLoadProgressEvent,
                cancellationToken: _loadCts.Token);

            if (result.IsSuccess && result.Value != null)
            {
                await RunOnUiThreadAsync(() =>
                {
                    foreach (var candidate in result.Value.Candidates)
                    {
                        Candidates.Add(candidate);
                    }

                    TotalCount = result.Value.TotalCount;
                    HasMore = result.Value.HasMore;
                    _currentSkip = Candidates.Count;
                });
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? "Failed to load more member users";
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

    private GetMailboxProvisioningCandidatesRequest BuildRequest(int skip, int? pageSize = null)
        => new()
        {
            SearchQuery = string.IsNullOrWhiteSpace(SearchQuery) ? null : SearchQuery.Trim(),
            OnlyWithoutLicense = OnlyWithoutLicense,
            OnlyWithoutMail = OnlyWithoutMail,
            Skip = skip,
            PageSize = pageSize ?? PageSize
        };

    private MailboxProvisioningCandidateDto? ResolveSelection(string? userPrincipalName)
    {
        if (string.IsNullOrWhiteSpace(userPrincipalName))
        {
            return Candidates.FirstOrDefault();
        }

        return Candidates.FirstOrDefault(candidate =>
                   string.Equals(candidate.UserPrincipalName, userPrincipalName, StringComparison.OrdinalIgnoreCase))
               ?? Candidates.FirstOrDefault();
    }

    private async Task LoadSelectedCandidateLicensesAsync()
    {
        if (SelectedCandidate == null)
        {
            Licenses.Reset();
            return;
        }

        Licenses.LoadAssignedLicenses(SelectedCandidate.AssignedLicenses);

        try
        {
            await Licenses.RefreshAsync(CancellationToken.None);
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Warning, $"Mailbox provisioning licenses refresh failed: {ex.Message}");
        }
    }

    private async Task RefreshAfterLicenseMutationAsync(CancellationToken cancellationToken)
    {
        var selectedUpn = SelectedCandidate?.UserPrincipalName;
        await RefreshAsync(cancellationToken);

        if (string.IsNullOrWhiteSpace(selectedUpn))
        {
            return;
        }

        SelectedCandidate = Candidates.FirstOrDefault(candidate =>
            string.Equals(candidate.UserPrincipalName, selectedUpn, StringComparison.OrdinalIgnoreCase));
    }

    private void OnLicensesPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(MailboxLicensesViewModel.IsLicenseSaving))
        {
            OnPropertyChanged(nameof(IsLicenseSaving));
        }
    }

    private void OnShellPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ShellViewModel.IsExchangeConnected) && !_shellViewModel.IsExchangeConnected)
        {
            Reset();
        }
    }

    private void HandleLoadProgressEvent(EventEnvelope evt)
    {
        _ = RunOnUiThreadAsync(() => LoadProgress.Apply(evt));
    }

    private static int GetRefreshPageSize(int loadedCount)
        => Math.Max(PageSize, loadedCount);
}
