using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Paging;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

public partial class MobileDevicesViewModel : ViewModelBase
{
    private const NavigationPage AlertPage = NavigationPage.MobileDevices;
    private readonly IWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;
    private readonly DebounceHelper _searchDebounce = new();

    private bool _isLoading;
    private bool _isLoadingSelection;
    private bool _isApplyingAction;
    private string? _errorMessage;
    private string? _capabilityMessage;
    private string? _searchQuery;
    private string _accessStateFilter = "All";
    private int _totalCount;
    private bool _isTotalCountExact = true;
    private int _currentSkip;
    private const int PageSize = PagingDefaults.DefaultPageSize;
    private bool _hasMore;
    private MobileDeviceListItemDto? _selectedDevice;
    private string? _selectedMailboxPolicyIdentity;
    private double _loadingProgress;
    private string? _loadingStatus;
    private int? _loadingCurrentItem;
    private int? _loadingTotalItems;
    private bool _suppressSelectedDeviceLoad;

    public MobileDevicesViewModel(IWorkerService workerService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;

        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
        LoadMoreCommand = new AsyncRelayCommand(LoadMoreAsync, () => CanLoadMore);
        AllowDeviceCommand = new AsyncRelayCommand(() => SetAccessStateAsync("Allowed"), () => CanManageDeviceAccessState);
        BlockDeviceCommand = new AsyncRelayCommand(() => SetAccessStateAsync("Blocked"), () => CanManageDeviceAccessState);
        QuarantineDeviceCommand = new AsyncRelayCommand(() => SetAccessStateAsync("Quarantined"), () => CanManageDeviceAccessState);
        RemoteWipeCommand = new AsyncRelayCommand(RemoteWipeAsync, () => CanRemoteWipe);
        AssignPolicyCommand = new AsyncRelayCommand(AssignPolicyAsync, () => CanAssignPolicy);
        _shellViewModel.PropertyChanged += OnShellViewModelPropertyChanged;
    }

    public ObservableCollection<MobileDeviceListItemDto> Devices { get; } = new();
    public ObservableCollection<MobileDeviceMailboxPolicyDto> Policies { get; } = new();

    public IReadOnlyList<string> AccessStateFilters { get; } = new[]
    {
        "All",
        "Allowed",
        "Blocked",
        "Quarantined",
        "DeviceDiscovery",
        "Unknown"
    };

    public bool IsLoading
    {
        get => _isLoading;
        private set
        {
            if (SetProperty(ref _isLoading, value))
            {
                OnPropertyChanged(nameof(IsBusy));
                OnPropertyChanged(nameof(LoadingOverlayText));
                RaiseCanExecuteChanged();
            }
        }
    }

    public bool IsApplyingAction
    {
        get => _isApplyingAction;
        private set
        {
            if (SetProperty(ref _isApplyingAction, value))
            {
                RaiseCanExecuteChanged();
            }
        }
    }

    public bool IsLoadingSelection
    {
        get => _isLoadingSelection;
        private set
        {
            if (SetProperty(ref _isLoadingSelection, value))
            {
                OnPropertyChanged(nameof(IsBusy));
                OnPropertyChanged(nameof(LoadingOverlayText));
                RaiseCanExecuteChanged();
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

    public bool HasError => !string.IsNullOrWhiteSpace(ErrorMessage);

    public bool IsBusy => IsLoading || IsLoadingSelection;

    public string? CapabilityMessage
    {
        get => _capabilityMessage;
        private set
        {
            if (SetProperty(ref _capabilityMessage, value))
            {
                OnPropertyChanged(nameof(HasCapabilityMessage));
            }
        }
    }

    public bool HasCapabilityMessage => !string.IsNullOrWhiteSpace(CapabilityMessage);

    public double LoadingProgress
    {
        get => _loadingProgress;
        private set
        {
            if (SetProperty(ref _loadingProgress, value))
            {
                OnPropertyChanged(nameof(LoadingPercentText));
            }
        }
    }

    public string? LoadingStatus
    {
        get => _loadingStatus;
        private set
        {
            if (SetProperty(ref _loadingStatus, value))
            {
                OnPropertyChanged(nameof(LoadingOverlayText));
            }
        }
    }

    public int? LoadingCurrentItem
    {
        get => _loadingCurrentItem;
        private set
        {
            if (SetProperty(ref _loadingCurrentItem, value))
            {
                OnPropertyChanged(nameof(HasLoadingCount));
                OnPropertyChanged(nameof(LoadingCountText));
            }
        }
    }

    public int? LoadingTotalItems
    {
        get => _loadingTotalItems;
        private set
        {
            if (SetProperty(ref _loadingTotalItems, value))
            {
                OnPropertyChanged(nameof(LoadingCountText));
            }
        }
    }

    public bool HasLoadingCount => LoadingCurrentItem.HasValue;

    public string LoadingPercentText => FormatProgressPercent(LoadingProgress);

    public string? LoadingCountText => FormatProgressCount(LoadingCurrentItem, LoadingTotalItems, "items");

    public string LoadingOverlayText => LoadingStatus
        ?? (IsLoading ? "Loading mobile devices list..." : "Loading device details...");

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

    public string AccessStateFilter
    {
        get => _accessStateFilter;
        set
        {
            if (SetProperty(ref _accessStateFilter, value))
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
                RaiseCanExecuteChanged();
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

    public string StatusText => IsMobileDevicesFeatureAvailable
        ? IsTotalCountExact
            ? $"{Devices.Count} of {TotalCount} device"
            : $"{Devices.Count}+ device loaded"
        : "Module not available";

    public MobileDeviceListItemDto? SelectedDevice
    {
        get => _selectedDevice;
        set
        {
            if (SetProperty(ref _selectedDevice, value))
            {
                SelectedMailboxPolicyIdentity = ResolveSelectedPolicy(value);
                RaiseDeviceSelectionChanged();

                if (!_suppressSelectedDeviceLoad && value != null)
                {
                    _ = SafeLoadSelectedDeviceAsync(value);
                }
            }
        }
    }

    public string? SelectedMailboxPolicyIdentity
    {
        get => _selectedMailboxPolicyIdentity;
        set
        {
            if (SetProperty(ref _selectedMailboxPolicyIdentity, value))
            {
                RaiseCanExecuteChanged();
            }
        }
    }

    public string SelectedDeviceSummary => SelectedDevice == null
        ? "Select a device to review details and available actions."
        : $"{SelectedDevice.UserDisplayName ?? SelectedDevice.UserPrincipalName ?? SelectedDevice.MailboxIdentity} - {SelectedDevice.DeviceType ?? "Unknown"} / {SelectedDevice.DeviceId}";

    public bool IsMobileDevicesFeatureAvailable => GetCapabilityState().IsModuleAvailable;
    public bool CanLoadPolicies => GetCapabilityState().CanLoadPolicies;
    public bool CanManageDeviceAccessState => !IsBusy && !IsApplyingAction && _shellViewModel.IsExchangeConnected && SelectedDevice != null && GetCapabilityState().CanManageAccessState;
    public bool CanRemoteWipe => !IsBusy && !IsApplyingAction && _shellViewModel.IsExchangeConnected && SelectedDevice != null && GetCapabilityState().CanRemoteWipe;
    public bool CanRefresh => !IsBusy && !IsApplyingAction && _shellViewModel.IsExchangeConnected && IsMobileDevicesFeatureAvailable;
    public bool CanLoadMore => !IsBusy && !IsApplyingAction && HasMore && _shellViewModel.IsExchangeConnected && IsMobileDevicesFeatureAvailable;
    public bool CanManageSelectedDevice => CanManageDeviceAccessState || CanRemoteWipe || CanAssignPolicy;
    public bool CanAssignPolicy => !IsBusy
        && !IsApplyingAction
        && _shellViewModel.IsExchangeConnected
        && SelectedDevice != null
        && !string.IsNullOrWhiteSpace(SelectedDevice.MailboxIdentity)
        && GetCapabilityState().CanAssignPolicy;

    public ICommand RefreshCommand { get; }
    public ICommand LoadMoreCommand { get; }
    public ICommand AllowDeviceCommand { get; }
    public ICommand BlockDeviceCommand { get; }
    public ICommand QuarantineDeviceCommand { get; }
    public ICommand RemoteWipeCommand { get; }
    public ICommand AssignPolicyCommand { get; }

    public async Task LoadAsync()
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            CapabilityMessage = null;
            Devices.Clear();
            Policies.Clear();
            ErrorMessage = null;
            TotalCount = 0;
            IsTotalCountExact = true;
            HasMore = false;
            IsLoadingSelection = false;
            ClearLoadingProgress();
            _shellViewModel.ClearPageAlert(AlertPage);
            return;
        }

        if (!await EnsureCapabilityStateAsync(CancellationToken.None))
        {
            return;
        }

        if (Devices.Count == 0)
        {
            await RefreshAsync(CancellationToken.None);
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
            _shellViewModel.AddLog(LogLevel.Error, $"Mobile devices refresh failed: {ex.Message}", "MobileDevices");
        }
    }

    private async Task SafeLoadSelectedDeviceAsync(MobileDeviceListItemDto selected)
    {
        try
        {
            await LoadSelectedDeviceAsync(selected, CancellationToken.None);
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Mobile device detail load failed: {ex.Message}", "MobileDevices");
        }
    }

    private GetMobileDevicesRequest BuildRequest(int skip, int? pageSize = null)
    {
        return new GetMobileDevicesRequest
        {
            SearchQuery = string.IsNullOrWhiteSpace(SearchQuery) ? null : SearchQuery.Trim(),
            AccessState = NormalizeAccessStateFilter(AccessStateFilter),
            PageSize = pageSize ?? PageSize,
            Skip = skip,
            SortBy = "UserDisplayName"
        };
    }

    private string? ResolveSelectedPolicy(MobileDeviceListItemDto? device)
    {
        if (device == null)
        {
            return null;
        }

        var byIdentity = Policies.FirstOrDefault(policy =>
            string.Equals(policy.Identity, device.CurrentMailboxPolicy, StringComparison.OrdinalIgnoreCase));
        if (byIdentity != null)
        {
            return byIdentity.Identity;
        }

        var byName = Policies.FirstOrDefault(policy =>
            string.Equals(policy.Name, device.CurrentMailboxPolicy, StringComparison.OrdinalIgnoreCase));
        return byName?.Identity;
    }

    private void RaiseDeviceSelectionChanged()
    {
        OnPropertyChanged(nameof(SelectedDeviceSummary));
        RaiseCanExecuteChanged();
    }

    private bool IsStillSelected(string identity)
        => SelectedDevice != null && string.Equals(SelectedDevice.Identity, identity, StringComparison.OrdinalIgnoreCase);

    private void RaiseCanExecuteChanged()
    {
        OnPropertyChanged(nameof(IsMobileDevicesFeatureAvailable));
        OnPropertyChanged(nameof(CanLoadPolicies));
        OnPropertyChanged(nameof(CanRefresh));
        OnPropertyChanged(nameof(CanLoadMore));
        OnPropertyChanged(nameof(CanManageDeviceAccessState));
        OnPropertyChanged(nameof(CanRemoteWipe));
        OnPropertyChanged(nameof(CanManageSelectedDevice));
        OnPropertyChanged(nameof(CanAssignPolicy));
        OnPropertyChanged(nameof(StatusText));
        CommandManager.InvalidateRequerySuggested();
    }

    private async Task<bool> EnsureCapabilityStateAsync(CancellationToken cancellationToken)
    {
        CapabilityMapDto? capabilities = _shellViewModel.Capabilities;

        if (_shellViewModel.Capabilities == null)
        {
            var capabilitiesResult = await _workerService.DetectCapabilitiesAsync(cancellationToken: cancellationToken);
            if (!capabilitiesResult.IsSuccess || capabilitiesResult.Value == null)
            {
                Devices.Clear();
                Policies.Clear();
                SelectedDevice = null;
                SelectedMailboxPolicyIdentity = null;
                TotalCount = 0;
                IsTotalCountExact = true;
                HasMore = false;
                IsLoadingSelection = false;
                ClearLoadingProgress();
                CapabilityMessage = "Unable to verify Mobile Devices capabilities for the current Exchange session. The module will not be loaded automatically.";
                ErrorMessage = null;
                RaiseCanExecuteChanged();
                return false;
            }

            capabilities = capabilitiesResult.Value;
        }

        var capabilityState = GetCapabilityState(capabilities);
        CapabilityMessage = capabilityState.Message;

        if (capabilityState.IsModuleAvailable)
        {
            return true;
        }

        Devices.Clear();
        Policies.Clear();
        SelectedDevice = null;
        SelectedMailboxPolicyIdentity = null;
        TotalCount = 0;
        IsTotalCountExact = true;
        HasMore = false;
        IsLoadingSelection = false;
        ClearLoadingProgress();
        ErrorMessage = null;
        RaiseCanExecuteChanged();
        return false;
    }

    private MobileDevicesCapabilityState GetCapabilityState(CapabilityMapDto? capabilities = null) => MobileDevicesCapabilityState.From(capabilities ?? _shellViewModel.Capabilities);

    private void OnShellViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ShellViewModel.Capabilities) ||
            e.PropertyName == nameof(ShellViewModel.IsExchangeConnected))
        {
            CapabilityMessage = _shellViewModel.IsExchangeConnected
                ? GetCapabilityState().Message
                : null;

            if (!_shellViewModel.IsExchangeConnected)
            {
                IsLoadingSelection = false;
                ClearLoadingProgress();
                _shellViewModel.ClearPageAlert(AlertPage);
            }

            RaiseCanExecuteChanged();
        }
    }

    private void HandleWorkerEvent(EventEnvelope evt)
    {
        if (evt.EventType != EventType.Progress)
        {
            return;
        }

        var progress = JsonMessageSerializer.ExtractPayload<ProgressEventPayload>(evt.Payload);
        if (progress == null)
        {
            return;
        }

        LoadingStatus = progress.StatusMessage;
        LoadingProgress = progress.PercentComplete;
        LoadingCurrentItem = progress.CurrentItem;
        LoadingTotalItems = progress.TotalItems;
    }

    private static string? NormalizeAccessStateFilter(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || string.Equals(value, "All", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return value;
    }

    private static int GetRefreshPageSize(int loadedCount)
        => Math.Max(PageSize, loadedCount);

    private bool HasWorkspaceData => Devices.Count > 0;
}


