using System.Collections.ObjectModel;
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
    private string? _newMailboxPrimarySmtpAddress;
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

    public string? NewMailboxPrimarySmtpAddress
    {
        get => _newMailboxPrimarySmtpAddress;
        set
        {
            if (SetProperty(ref _newMailboxPrimarySmtpAddress, value))
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

    public bool CanCreateMailbox =>
        !IsLoading && !IsCreatingMailbox && _shellViewModel.IsExchangeConnected &&
        !string.IsNullOrWhiteSpace(NewMailboxDisplayName) &&
        !string.IsNullOrWhiteSpace(NewMailboxAlias) &&
        !string.IsNullOrWhiteSpace(NewMailboxPrimarySmtpAddress);

    public bool CanRefresh => !IsLoading && _shellViewModel.IsExchangeConnected;
    public bool CanLoadMore => !IsLoading && HasMore && _shellViewModel.IsExchangeConnected;

    public string StatusText => $"{Mailboxes.Count} of {TotalCount} mailboxes";

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
                PrimarySmtpAddress = NewMailboxPrimarySmtpAddress!.Trim(),
                MailboxType = "Shared"
            };

            _shellViewModel.AddLog(LogLevel.Information, $"Creating mailbox {request.PrimarySmtpAddress}...");

            var result = await _workerService.CreateMailboxAsync(request, cancellationToken: cancellationToken);
            if (result.IsSuccess)
            {
                _shellViewModel.AddLog(LogLevel.Information, "Mailbox created successfully");
                NewMailboxDisplayName = string.Empty;
                NewMailboxAlias = string.Empty;
                NewMailboxPrimarySmtpAddress = string.Empty;
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
