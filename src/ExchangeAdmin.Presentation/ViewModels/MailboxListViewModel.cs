using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

public class MailboxListViewModel : ViewModelBase
{
    private readonly IWorkerService _workerService;
    private readonly NavigationService _navigationService;
    private readonly ShellViewModel _shellViewModel;

    private readonly DebounceHelper _searchDebounce = new();
    private CancellationTokenSource? _loadCts;

    private bool _isLoading;
    private string? _errorMessage;
    private string? _searchQuery;
    private string _recipientTypeFilter = "UserMailbox";

    private int _totalCount;
    private int _currentSkip;
    private const int PageSize = 50;
    private bool _hasMore;

    private MailboxListItemDto? _selectedMailbox;

    private string? _newMailboxDisplayName;
    private string? _newMailboxAlias;
    private string? _newMailboxLocalPart;
    private string? _selectedMailboxDomain;
    private string? _newMailboxPassword;
    private bool _isCreatingMailbox;

    public MailboxListViewModel(IWorkerService workerService, NavigationService navigationService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _navigationService = navigationService;
        _shellViewModel = shellViewModel;

        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
        LoadMoreCommand = new AsyncRelayCommand(LoadMoreAsync, () => CanLoadMore);
        CancelCommand = new RelayCommand(Cancel, () => IsLoading);
        ViewDetailsCommand = new RelayCommand<MailboxListItemDto>(ViewDetails, m => m != null);
        CreateMailboxCommand = new AsyncRelayCommand(CreateMailboxAsync, () => CanCreateMailbox);
    }

    #region Properties

    public ObservableCollection<MailboxListItemDto> Mailboxes { get; } = new();
    public ObservableCollection<string> AvailableMailDomains { get; } = new();

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
                // Use DebounceHelper for more efficient debouncing
                _searchDebounce.Debounce(SafeRefreshAsync, 300);
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
                SafeRefreshAsync();
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

    public string? NewMailboxPassword
    {
        get => _newMailboxPassword;
        set
        {
            if (SetProperty(ref _newMailboxPassword, value))
            {
                OnPropertyChanged(nameof(CanCreateMailbox));
                CommandManager.InvalidateRequerySuggested();
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
        (IsSharedMailboxCreation || !string.IsNullOrWhiteSpace(NewMailboxPassword));

    public bool CanRefresh => !IsLoading && _shellViewModel.IsExchangeConnected;
    public bool CanLoadMore => !IsLoading && HasMore && _shellViewModel.IsExchangeConnected;

    public string StatusText => $"{Mailboxes.Count} of {TotalCount} mailboxes";

    public string CreateMailboxSectionTitle => IsSharedMailboxCreation ? "New Shared Mailbox" : "New Mailbox";

    #endregion

    #region Commands

    public ICommand RefreshCommand { get; }
    public ICommand LoadMoreCommand { get; }
    public ICommand CancelCommand { get; }
    public ICommand ViewDetailsCommand { get; }
    public ICommand CreateMailboxCommand { get; }

    #endregion

    #region Methods

    private async void SafeRefreshAsync()
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
            await RunOnUiThreadAsync(() => Mailboxes.Clear());
            ErrorMessage = "Not connected to Exchange Online";
            return;
        }

        await LoadAcceptedDomainsAsync(CancellationToken.None);
        await RefreshAsync(CancellationToken.None);
    }

    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        _loadCts?.Cancel();
        if (!_shellViewModel.IsExchangeConnected)
        {
            IsLoading = false;
            await RunOnUiThreadAsync(() => Mailboxes.Clear());
            ErrorMessage = "Not connected to Exchange Online";
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
            var request = new GetMailboxesRequest
            {
                RecipientTypeDetails = _recipientTypeFilter,
                SearchQuery = _searchQuery,
                PageSize = PageSize,
                Skip = 0
            };

            var result = await _workerService.GetMailboxesAsync(
                request,
                eventHandler: null,
                cancellationToken: _loadCts.Token);

            if (result.IsSuccess && result.Value != null)
            {
                await RunOnUiThreadAsync(() =>
                {
                    // Use ReplaceAll for smoother UI updates instead of Clear + Add
                    Mailboxes.ReplaceAll(result.Value.Mailboxes);
                    TotalCount = result.Value.TotalCount;
                    HasMore = result.Value.HasMore;
                    _currentSkip = result.Value.Mailboxes.Count;
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
        ErrorMessage = null;

        try
        {
            var request = new GetMailboxesRequest
            {
                RecipientTypeDetails = _recipientTypeFilter,
                SearchQuery = _searchQuery,
                PageSize = PageSize,
                Skip = _currentSkip
            };

            var result = await _workerService.GetMailboxesAsync(
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
            var request = new CreateMailboxRequest
            {
                DisplayName = NewMailboxDisplayName!.Trim(),
                Alias = NewMailboxAlias!.Trim(),
                PrimarySmtpAddress = BuildPrimarySmtpAddress(),
                MailboxType = IsSharedMailboxCreation ? "Shared" : "User",
                Password = IsSharedMailboxCreation ? null : NewMailboxPassword?.Trim()
            };

            _shellViewModel.AddLog(LogLevel.Information, $"Creating mailbox {request.PrimarySmtpAddress}...");

            var result = await _workerService.CreateMailboxAsync(request, cancellationToken: cancellationToken);
            if (result.IsSuccess)
            {
                _shellViewModel.AddLog(LogLevel.Information, "Mailbox created successfully");
                NewMailboxDisplayName = string.Empty;
                NewMailboxAlias = string.Empty;
                NewMailboxLocalPart = string.Empty;
                NewMailboxPassword = string.Empty;
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

    private async Task LoadAcceptedDomainsAsync(CancellationToken cancellationToken)
    {
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

    private void ViewDetails(MailboxListItemDto? mailbox)
    {
        if (mailbox == null) return;
        _navigationService.NavigateToDetails(_navigationService.CurrentPage, mailbox.Identity, mailbox);
    }

    public void Cancel()
    {
        _loadCts?.Cancel();
    }

    #endregion
}
