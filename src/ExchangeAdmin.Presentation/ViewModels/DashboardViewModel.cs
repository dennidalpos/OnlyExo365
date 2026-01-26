using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;




public class DashboardViewModel : ViewModelBase
{
    private readonly IWorkerService _workerService;
    private readonly NavigationService _navigationService;
    private readonly ShellViewModel _shellViewModel;

    private CancellationTokenSource? _loadCts;

    private bool _isLoading;
    private DashboardStatsDto? _stats;
    private string? _errorMessage;

    public DashboardViewModel(IWorkerService workerService, NavigationService navigationService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _navigationService = navigationService;
        _shellViewModel = shellViewModel;

        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
        NavigateToMailboxesCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.Mailboxes));
        NavigateToSharedMailboxesCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.SharedMailboxes));
        NavigateToDistributionListsCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.DistributionLists));
    }

    #region Properties

    public bool IsLoading
    {
        get => _isLoading;
        private set
        {
            if (SetProperty(ref _isLoading, value))
            {
                OnPropertyChanged(nameof(CanRefresh));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public DashboardStatsDto? Stats
    {
        get => _stats;
        private set
        {
            if (SetProperty(ref _stats, value))
            {
                OnPropertyChanged(nameof(HasStats));
                OnPropertyChanged(nameof(UserMailboxCount));
                OnPropertyChanged(nameof(SharedMailboxCount));
                OnPropertyChanged(nameof(RoomMailboxCount));
                OnPropertyChanged(nameof(EquipmentMailboxCount));
                OnPropertyChanged(nameof(TotalMailboxCount));
                OnPropertyChanged(nameof(DistributionGroupCount));
                OnPropertyChanged(nameof(DynamicDistributionGroupCount));
                OnPropertyChanged(nameof(UnifiedGroupCount));
                OnPropertyChanged(nameof(TotalGroupCount));
                OnPropertyChanged(nameof(IsLargeTenant));
                OnPropertyChanged(nameof(Warnings));
                OnPropertyChanged(nameof(HasWarnings));
                OnPropertyChanged(nameof(ShowUnifiedGroups));
                OnPropertyChanged(nameof(IsApproximate));
            }
        }
    }

    public bool HasStats => Stats != null;

    public int UserMailboxCount => Stats?.MailboxCounts.UserMailboxes ?? 0;
    public int SharedMailboxCount => Stats?.MailboxCounts.SharedMailboxes ?? 0;
    public int RoomMailboxCount => Stats?.MailboxCounts.RoomMailboxes ?? 0;
    public int EquipmentMailboxCount => Stats?.MailboxCounts.EquipmentMailboxes ?? 0;
    public int TotalMailboxCount => Stats?.MailboxCounts.Total ?? 0;

    public int DistributionGroupCount => Stats?.GroupCounts.DistributionGroups ?? 0;
    public int DynamicDistributionGroupCount => Stats?.GroupCounts.DynamicDistributionGroups ?? 0;
    public int? UnifiedGroupCount => Stats?.GroupCounts.UnifiedGroups;
    public int TotalGroupCount => Stats?.GroupCounts.Total ?? 0;

    public bool IsLargeTenant => Stats?.IsLargeTenant ?? false;
    public List<string> Warnings => Stats?.Warnings ?? new List<string>();
    public bool HasWarnings => Warnings.Count > 0;
    public bool ShowUnifiedGroups => Stats?.GroupCounts.UnifiedGroupsAvailable ?? false;
    public bool IsApproximate => Stats?.MailboxCounts.IsApproximate ?? false;

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

    public bool CanRefresh => !IsLoading && _shellViewModel.IsExchangeConnected;

    #endregion

    #region Commands

    public ICommand RefreshCommand { get; }
    public ICommand NavigateToMailboxesCommand { get; }
    public ICommand NavigateToSharedMailboxesCommand { get; }
    public ICommand NavigateToDistributionListsCommand { get; }

    #endregion

    #region Methods

    public async Task LoadAsync()
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            Stats = null;
            ErrorMessage = "Not connected to Exchange Online";
            return;
        }

        await RefreshAsync(CancellationToken.None);
    }

    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        _loadCts?.Cancel();
        _loadCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

        IsLoading = true;
        ErrorMessage = null;

        try
        {
            Console.WriteLine("[DashboardViewModel] Requesting dashboard stats...");
            var result = await _workerService.GetDashboardStatsAsync(
                new GetDashboardStatsRequest { IncludeUnifiedGroups = true },
                eventHandler: null,
                cancellationToken: _loadCts.Token);

            Console.WriteLine($"[DashboardViewModel] Response received - IsSuccess: {result.IsSuccess}, HasValue: {result.Value != null}");

            if (result.IsSuccess && result.Value != null)
            {
                Console.WriteLine($"[DashboardViewModel] Stats loaded - Mailboxes: {result.Value.MailboxCounts.Total}, Groups: {result.Value.GroupCounts.Total}");
                Stats = result.Value;

                
                foreach (var warning in result.Value.Warnings)
                {
                    _shellViewModel.AddLog(LogLevel.Warning, warning);
                }
            }
            else if (result.WasCancelled)
            {
                
            }
            else
            {
                var errorDetails = result.Error != null
                    ? $"{result.Error.Code}: {result.Error.Message}"
                    : "Failed to load dashboard (no error details)";
                ErrorMessage = errorDetails;
                _shellViewModel.AddLog(LogLevel.Error, $"Dashboard load failed: {errorDetails}");

                
                Console.WriteLine($"[DashboardViewModel] Error: {errorDetails}");
                if (result.Error != null)
                {
                    Console.WriteLine($"[DashboardViewModel] ErrorCode: {result.Error.Code}, IsTransient: {result.Error.IsTransient}");
                }
            }
        }
        catch (OperationCanceledException)
        {
            
        }
        catch (Exception ex)
        {
            ErrorMessage = $"Exception: {ex.GetType().Name} - {ex.Message}";
            _shellViewModel.AddLog(LogLevel.Error, $"Dashboard exception: {ex.GetType().Name} - {ex.Message}");
            Console.WriteLine($"[DashboardViewModel] Exception: {ex}");
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

    #endregion
}
