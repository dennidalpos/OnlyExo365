using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Paging;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

public class MailboxListViewModel : ViewModelBase
{
    private readonly IMailboxesWorkerService _workerService;
    private readonly NavigationService _navigationService;
    private readonly ShellViewModel _shellViewModel;

    private readonly DebounceHelper _searchDebounce = new();
    private CancellationTokenSource? _loadCts;

    private bool _isLoading;
    private string? _errorMessage;
    private string? _searchQuery;
    private string _recipientTypeFilter = "UserMailbox";
    private bool _isProvisioningWorkspace;

    private int _totalCount;
    private bool _isTotalCountExact = true;
    private int _currentSkip;
    private const int PageSize = PagingDefaults.DefaultPageSize;
    private bool _hasMore;

    private MailboxListItemDto? _selectedMailbox;

    private string? _newMailboxDisplayName;
    private string? _newMailboxAlias;
    private string? _newMailboxLocalPart;
    private string? _selectedMailboxDomain;
    private string? _pendingNewMailboxPassword;
    private int _newMailboxPasswordClearTrigger;
    private bool _isCreatingMailbox;

    public MailboxListViewModel(IMailboxesWorkerService workerService, NavigationService navigationService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _navigationService = navigationService;
        _shellViewModel = shellViewModel;
        Provisioning = new MailboxProvisioningCandidatesViewModel(workerService, shellViewModel);
        _shellViewModel.PropertyChanged += OnShellPropertyChanged;
        Provisioning.PropertyChanged += OnProvisioningPropertyChanged;

        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
        LoadMoreCommand = new AsyncRelayCommand(LoadMoreAsync, () => CanLoadMore);
        CancelCommand = new RelayCommand(Cancel, () => IsLoading);
        ViewDetailsCommand = new RelayCommand<MailboxListItemDto>(ViewDetails, m => m != null);
        CreateMailboxCommand = new AsyncRelayCommand(CreateMailboxAsync, () => CanCreateMailbox);
        ShowMailboxInventoryCommand = new AsyncRelayCommand(ShowMailboxInventoryAsync);
        ShowProvisioningWorkspaceCommand = new AsyncRelayCommand(ShowProvisioningWorkspaceAsync);
    }

    #region Properties

    public ObservableCollection<MailboxListItemDto> Mailboxes { get; } = new();
    public ObservableCollection<string> AvailableMailDomains { get; } = new();
    public MailboxProvisioningCandidatesViewModel Provisioning { get; }
    public CollectionLoadProgressViewModel LoadProgress { get; } = new();

    public bool IsMailboxInventoryWorkspace => !_isProvisioningWorkspace;

    public bool IsProvisioningWorkspace
    {
        get => _isProvisioningWorkspace;
        private set
        {
            if (SetProperty(ref _isProvisioningWorkspace, value))
            {
                OnPropertyChanged(nameof(IsMailboxInventoryWorkspace));
            }
        }
    }

    public bool IsProvisioningLoading => Provisioning.IsLoading;

    public bool IsAssigningProvisioningLicense => Provisioning.IsLicenseSaving;

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

    public string RecipientTypeFilter
    {
        get => _recipientTypeFilter;
        set
        {
            if (SetProperty(ref _recipientTypeFilter, value))
            {
                TriggerRefreshFromUi();
                OnPropertyChanged(nameof(CreateMailboxSectionTitle));
                OnPropertyChanged(nameof(IsSharedMailboxCreation));
                OnPropertyChanged(nameof(IsUserMailboxCreation));
                OnPropertyChanged(nameof(CanCreateMailbox));
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

    public MailboxListItemDto? SelectedMailbox
    {
        get => _selectedMailbox;
        set
        {
            if (SetProperty(ref _selectedMailbox, value) && value != null)
            {
                ViewDetails(value);
            }
        }
    }


    public string? NewMailboxDisplayName
    {
        get => _newMailboxDisplayName;
        set
        {
            if (SetProperty(ref _newMailboxDisplayName, value))
            {
                OnPropertyChanged(nameof(CanCreateMailbox));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string? NewMailboxAlias
    {
        get => _newMailboxAlias;
        set
        {
            if (SetProperty(ref _newMailboxAlias, value))
            {
                OnPropertyChanged(nameof(CanCreateMailbox));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string? NewMailboxLocalPart
    {
        get => _newMailboxLocalPart;
        set
        {
            if (SetProperty(ref _newMailboxLocalPart, value))
            {
                OnPropertyChanged(nameof(CanCreateMailbox));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string? SelectedMailboxDomain
    {
        get => _selectedMailboxDomain;
        set
        {
            if (SetProperty(ref _selectedMailboxDomain, value))
            {
                OnPropertyChanged(nameof(CanCreateMailbox));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool HasNewMailboxPassword
    {
        get => !string.IsNullOrWhiteSpace(_pendingNewMailboxPassword);
    }

    public int NewMailboxPasswordClearTrigger
    {
        get => _newMailboxPasswordClearTrigger;
        private set
        {
            if (SetProperty(ref _newMailboxPasswordClearTrigger, value))
            {
            }
        }
    }

    public bool IsCreatingMailbox
    {
        get => _isCreatingMailbox;
        private set
        {
            if (SetProperty(ref _isCreatingMailbox, value))
            {
                OnPropertyChanged(nameof(CanCreateMailbox));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool IsSharedMailboxCreation => string.Equals(RecipientTypeFilter, "SharedMailbox", StringComparison.OrdinalIgnoreCase);
    public bool IsUserMailboxCreation => !IsSharedMailboxCreation;

    public bool CanCreateMailbox =>
        !IsLoading && !IsCreatingMailbox && _shellViewModel.IsExchangeConnected &&
        !string.IsNullOrWhiteSpace(NewMailboxDisplayName) &&
        !string.IsNullOrWhiteSpace(NewMailboxAlias) &&
        !string.IsNullOrWhiteSpace(NewMailboxLocalPart) &&
        !string.IsNullOrWhiteSpace(SelectedMailboxDomain) &&
        (IsSharedMailboxCreation || HasNewMailboxPassword);

    public bool CanRefresh => !IsLoading && _shellViewModel.IsExchangeConnected;
    public bool CanLoadMore => !IsLoading && HasMore && _shellViewModel.IsExchangeConnected;

    public string StatusText => IsTotalCountExact
        ? $"{Mailboxes.Count} of {TotalCount} mailboxes"
        : $"{Mailboxes.Count}+ mailboxes loaded";

    public string CreateMailboxSectionTitle => IsSharedMailboxCreation ? "New Shared Mailbox" : "New Mailbox";

    #endregion

    #region Commands

    public ICommand RefreshCommand { get; }
    public ICommand LoadMoreCommand { get; }
    public ICommand CancelCommand { get; }
    public ICommand ViewDetailsCommand { get; }
    public ICommand CreateMailboxCommand { get; }
    public ICommand ShowMailboxInventoryCommand { get; }
    public ICommand ShowProvisioningWorkspaceCommand { get; }

    #endregion

    #region Methods

    public void SetRecipientTypeFilter(string? value, bool refresh = true)
    {
        var normalized = NormalizeRecipientTypeFilter(value) ?? "UserMailbox";
        if (string.Equals(_recipientTypeFilter, normalized, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        _recipientTypeFilter = normalized;
        OnPropertyChanged(nameof(RecipientTypeFilter));
        OnPropertyChanged(nameof(CreateMailboxSectionTitle));
        OnPropertyChanged(nameof(IsSharedMailboxCreation));
        OnPropertyChanged(nameof(IsUserMailboxCreation));
        if (string.Equals(normalized, "SharedMailbox", StringComparison.OrdinalIgnoreCase))
        {
            ClearNewMailboxPassword();
        }

        OnPropertyChanged(nameof(CanCreateMailbox));

        if (refresh)
        {
            TriggerRefreshFromUi();
        }
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
            _shellViewModel.AddLog(LogLevel.Error, $"Refresh failed: {ex.Message}");
        }
    }

    public async Task LoadAsync()
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            ClearNewMailboxPassword();
            await RunOnUiThreadAsync(() => Mailboxes.Clear());
            Provisioning.Reset();
            ErrorMessage = "Not connected to Exchange Online";
            return;
        }

        await LoadAcceptedDomainsAsync(CancellationToken.None);
        if (IsProvisioningWorkspace)
        {
            await Provisioning.LoadAsync(CancellationToken.None);
            return;
        }

        await RefreshAsync(CancellationToken.None);
    }

    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        _loadCts?.Cancel();
        if (!_shellViewModel.IsExchangeConnected)
        {
            IsLoading = false;
            ClearNewMailboxPassword();
            await RunOnUiThreadAsync(() => Mailboxes.Clear());
            ErrorMessage = "Not connected to Exchange Online";
            TotalCount = 0;
            IsTotalCountExact = true;
            HasMore = false;
            _currentSkip = 0;
            OnPropertyChanged(nameof(StatusText));
            return;
        }

        _loadCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

        IsLoading = true;
        LoadProgress.Start("Loading mailbox...", "mailbox");
        ErrorMessage = null;
        var refreshPageSize = GetRefreshPageSize(Mailboxes.Count);
        _currentSkip = 0;

        try
        {
            var request = new GetMailboxesRequest
            {
                RecipientTypeDetails = NormalizeRecipientTypeFilter(_recipientTypeFilter),
                SearchQuery = string.IsNullOrWhiteSpace(_searchQuery) ? null : _searchQuery.Trim(),
                PageSize = refreshPageSize,
                Skip = 0
            };

            var result = await _workerService.GetMailboxesAsync(
                request,
                eventHandler: HandleLoadProgressEvent,
                cancellationToken: _loadCts.Token);

            if (result.IsSuccess && result.Value != null)
            {
                await RunOnUiThreadAsync(() =>
                {
                    // Use ReplaceAll for smoother UI updates instead of Clear + Add
                    Mailboxes.ReplaceAll(result.Value.Mailboxes);
                    TotalCount = result.Value.TotalCount;
                    IsTotalCountExact = result.Value.IsTotalCountExact;
                    HasMore = result.Value.HasMore;
                    _currentSkip = Mailboxes.Count;
                    OnPropertyChanged(nameof(StatusText));
                });
            }
            else if (result.WasCancelled)
            {
            }
            else
            {
                var errorDetails = result.Error != null
                    ? $"{result.Error.Code}: {result.Error.Message}"
                    : "Failed to load mailboxes (no error details)";
                ErrorMessage = errorDetails;
                _shellViewModel.AddLog(LogLevel.Error, $"Mailbox load failed: {errorDetails}");
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            ErrorMessage = $"Exception: {ex.GetType().Name} - {ex.Message}";
            _shellViewModel.AddLog(LogLevel.Error, $"Mailbox exception: {ex.GetType().Name} - {ex.Message}");
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
            ErrorMessage = "Not connected to Exchange Online";
            return;
        }

        _loadCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

        IsLoading = true;
        LoadProgress.Start("Loading mailbox...", "mailbox");
        ErrorMessage = null;

        try
        {
            var request = new GetMailboxesRequest
            {
                RecipientTypeDetails = NormalizeRecipientTypeFilter(_recipientTypeFilter),
                SearchQuery = string.IsNullOrWhiteSpace(_searchQuery) ? null : _searchQuery.Trim(),
                PageSize = PageSize,
                Skip = _currentSkip
            };

            var result = await _workerService.GetMailboxesAsync(
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
                    IsTotalCountExact = result.Value.IsTotalCountExact;
                    HasMore = result.Value.HasMore;
                    _currentSkip = Mailboxes.Count;

                    OnPropertyChanged(nameof(StatusText));
                });
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? "Failed to load more mailboxes";
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


    private async Task CreateMailboxAsync(CancellationToken cancellationToken)
    {
        if (!CanCreateMailbox)
        {
            return;
        }

        IsCreatingMailbox = true;
        ErrorMessage = null;

        try
        {
            var password = IsSharedMailboxCreation ? null : CaptureNewMailboxPassword();
            var request = new CreateMailboxRequest
            {
                DisplayName = NewMailboxDisplayName!.Trim(),
                Alias = NewMailboxAlias!.Trim(),
                PrimarySmtpAddress = BuildPrimarySmtpAddress(),
                MailboxType = IsSharedMailboxCreation ? "Shared" : "User",
                Password = password
            };

            ClearNewMailboxPassword();

            _shellViewModel.AddLog(LogLevel.Information, $"Creating mailbox {request.PrimarySmtpAddress}...");

            var result = await _workerService.CreateMailboxAsync(request, cancellationToken: cancellationToken);
            if (result.IsSuccess)
            {
                _shellViewModel.AddLog(LogLevel.Information, "Mailbox created successfully");
                NewMailboxDisplayName = string.Empty;
                NewMailboxAlias = string.Empty;
                NewMailboxLocalPart = string.Empty;
                await RefreshAsync(cancellationToken);
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? "Failed to create mailbox";
                _shellViewModel.AddLog(LogLevel.Error, $"Mailbox creation failed: {ErrorMessage}");
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Mailbox creation error: {ex.Message}");
        }
        finally
        {
            IsCreatingMailbox = false;
        }
    }

    private string BuildPrimarySmtpAddress()
    {
        var localPart = NewMailboxLocalPart?.Trim() ?? string.Empty;
        var domain = SelectedMailboxDomain?.Trim() ?? string.Empty;
        return $"{localPart}@{domain}";
    }

    public void SetNewMailboxPassword(string? value)
    {
        var normalized = string.IsNullOrEmpty(value) ? null : value;
        if (string.Equals(_pendingNewMailboxPassword, normalized, StringComparison.Ordinal))
        {
            return;
        }

        _pendingNewMailboxPassword = normalized;
        OnPropertyChanged(nameof(HasNewMailboxPassword));
        OnPropertyChanged(nameof(CanCreateMailbox));
        CommandManager.InvalidateRequerySuggested();
    }

    private string? CaptureNewMailboxPassword()
    {
        return string.IsNullOrWhiteSpace(_pendingNewMailboxPassword)
            ? null
            : _pendingNewMailboxPassword.Trim();
    }

    private void ClearNewMailboxPassword()
    {
        var hadValue = !string.IsNullOrWhiteSpace(_pendingNewMailboxPassword);
        _pendingNewMailboxPassword = null;
        NewMailboxPasswordClearTrigger++;

        if (hadValue)
        {
            OnPropertyChanged(nameof(HasNewMailboxPassword));
            OnPropertyChanged(nameof(CanCreateMailbox));
            CommandManager.InvalidateRequerySuggested();
        }
    }

    private async Task LoadAcceptedDomainsAsync(CancellationToken cancellationToken)
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            await RunOnUiThreadAsync(() => AvailableMailDomains.Clear());
            if (!string.IsNullOrWhiteSpace(SelectedMailboxDomain))
            {
                SelectedMailboxDomain = null;
            }

            return;
        }

        try
        {
            var result = await _workerService.GetAcceptedDomainsAsync(new GetAcceptedDomainsRequest(), cancellationToken: cancellationToken);
            if (!result.IsSuccess || result.Value == null)
            {
                return;
            }

            var domains = result.Value.Domains
                .Select(d => d.DomainName?.Trim())
                .Where(d => !string.IsNullOrWhiteSpace(d))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(d => d, StringComparer.OrdinalIgnoreCase)
                .ToList();

            await RunOnUiThreadAsync(() =>
            {
                AvailableMailDomains.Clear();
                foreach (var domain in domains)
                {
                    AvailableMailDomains.Add(domain!);
                }
            });

            if (string.IsNullOrWhiteSpace(SelectedMailboxDomain))
            {
                SelectedMailboxDomain = result.Value.Domains.FirstOrDefault(d => d.Default)?.DomainName
                    ?? AvailableMailDomains.FirstOrDefault();
            }
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Warning, $"Unable to load accepted domains: {ex.Message}");
        }
    }


    private static string? NormalizeRecipientTypeFilter(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        if (string.Equals(value, "UserMailbox", StringComparison.OrdinalIgnoreCase))
        {
            return "UserMailbox";
        }

        if (string.Equals(value, "SharedMailbox", StringComparison.OrdinalIgnoreCase))
        {
            return "SharedMailbox";
        }

        if (string.Equals(value, "RoomMailbox", StringComparison.OrdinalIgnoreCase))
        {
            return "RoomMailbox";
        }

        if (string.Equals(value, "EquipmentMailbox", StringComparison.OrdinalIgnoreCase))
        {
            return "EquipmentMailbox";
        }

        return null;
    }

    private static int GetRefreshPageSize(int loadedCount)
        => Math.Max(PageSize, loadedCount);

    private void ViewDetails(MailboxListItemDto? mailbox)
    {
        if (mailbox == null) return;
        _navigationService.NavigateToDetails(_navigationService.CurrentPage, mailbox.Identity, mailbox);
    }

    public void Cancel()
    {
        _loadCts?.Cancel();
    }

    private void OnShellPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ShellViewModel.IsExchangeConnected) && !_shellViewModel.IsExchangeConnected)
        {
            ClearNewMailboxPassword();
        }
    }

    private async Task ShowMailboxInventoryAsync(CancellationToken cancellationToken)
    {
        if (IsMailboxInventoryWorkspace)
        {
            return;
        }

        IsProvisioningWorkspace = false;
        if (_shellViewModel.IsExchangeConnected)
        {
            await RefreshAsync(cancellationToken);
        }
    }

    private async Task ShowProvisioningWorkspaceAsync(CancellationToken cancellationToken)
    {
        if (IsProvisioningWorkspace)
        {
            if (_shellViewModel.IsExchangeConnected)
            {
                await Provisioning.LoadAsync(cancellationToken);
            }

            return;
        }

        IsProvisioningWorkspace = true;
        await Provisioning.LoadAsync(cancellationToken);
    }

    private void OnProvisioningPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(MailboxProvisioningCandidatesViewModel.IsLoading) or nameof(MailboxProvisioningCandidatesViewModel.IsLicenseSaving))
        {
            OnPropertyChanged(nameof(IsProvisioningLoading));
            OnPropertyChanged(nameof(IsAssigningProvisioningLicense));
        }
    }

    private void HandleLoadProgressEvent(EventEnvelope evt)
    {
        _ = RunOnUiThreadAsync(() => LoadProgress.Apply(evt));
    }

    #endregion
}
