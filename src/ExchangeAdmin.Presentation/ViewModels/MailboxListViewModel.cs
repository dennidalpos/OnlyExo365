using System.Collections.ObjectModel;
using System.Windows.Input;
using System.Windows.Threading;
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

    private readonly DispatcherTimer _searchDebounceTimer;
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

    public MailboxListViewModel(IWorkerService workerService, NavigationService navigationService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _navigationService = navigationService;
        _shellViewModel = shellViewModel;

                                        
        _searchDebounceTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(300)
        };
        _searchDebounceTimer.Tick += OnSearchDebounceElapsed;

                   
        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
        LoadMoreCommand = new AsyncRelayCommand(LoadMoreAsync, () => CanLoadMore);
        CancelCommand = new RelayCommand(Cancel, () => IsLoading);
        ViewDetailsCommand = new RelayCommand<MailboxListItemDto>(ViewDetails, m => m != null);
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
                                  
                _searchDebounceTimer.Stop();
                _searchDebounceTimer.Start();
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
                _ = RefreshAsync(CancellationToken.None);
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

    public bool CanRefresh => !IsLoading && _shellViewModel.IsExchangeConnected;
    public bool CanLoadMore => !IsLoading && HasMore && _shellViewModel.IsExchangeConnected;

    public string StatusText => $"{Mailboxes.Count} of {TotalCount} mailboxes";

    #endregion

    #region Commands

    public ICommand RefreshCommand { get; }
    public ICommand LoadMoreCommand { get; }
    public ICommand CancelCommand { get; }
    public ICommand ViewDetailsCommand { get; }

    #endregion

    #region Methods

    public async Task LoadAsync()
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            Mailboxes.Clear();
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
            Mailboxes.Clear();
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
        Mailboxes.Clear();

        try
        {
            var request = new GetMailboxesRequest
            {
                RecipientTypeDetails = _recipientTypeFilter,
                SearchQuery = _searchQuery,
                PageSize = PageSize,
                Skip = 0
            };

            Console.WriteLine($"[MailboxListViewModel] Requesting mailboxes...");
            var result = await _workerService.GetMailboxesAsync(
                request,
                eventHandler: null,
                cancellationToken: _loadCts.Token);

            Console.WriteLine($"[MailboxListViewModel] Response received - IsSuccess: {result.IsSuccess}, HasValue: {result.Value != null}");

            if (result.IsSuccess && result.Value != null)
            {
                Console.WriteLine($"[MailboxListViewModel] Mailboxes loaded - Count: {result.Value.Mailboxes.Count}, TotalCount: {result.Value.TotalCount}, HasMore: {result.Value.HasMore}");
                foreach (var mailbox in result.Value.Mailboxes)
                {
                    Mailboxes.Add(mailbox);
                }

                TotalCount = result.Value.TotalCount;
                HasMore = result.Value.HasMore;
                _currentSkip = result.Value.Mailboxes.Count;

                OnPropertyChanged(nameof(StatusText));
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
                Console.WriteLine($"[MailboxListViewModel] Error: {errorDetails}");
            }
        }
        catch (OperationCanceledException)
        {
                     
        }
        catch (Exception ex)
        {
            ErrorMessage = $"Exception: {ex.GetType().Name} - {ex.Message}";
            _shellViewModel.AddLog(LogLevel.Error, $"Mailbox exception: {ex.GetType().Name} - {ex.Message}");
            Console.WriteLine($"[MailboxListViewModel] Exception: {ex}");
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
                foreach (var mailbox in result.Value.Mailboxes)
                {
                    Mailboxes.Add(mailbox);
                }

                HasMore = result.Value.HasMore;
                _currentSkip += result.Value.Mailboxes.Count;

                OnPropertyChanged(nameof(StatusText));
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

    private void OnSearchDebounceElapsed(object? sender, EventArgs e)
    {
        _searchDebounceTimer.Stop();
        _ = RefreshAsync(CancellationToken.None);
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
