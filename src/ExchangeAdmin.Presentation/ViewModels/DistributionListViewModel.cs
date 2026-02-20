using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using System.Windows.Threading;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

             
                                                        
              
public class DistributionListViewModel : ViewModelBase
{
    private readonly IWorkerService _workerService;
    private readonly NavigationService _navigationService;
    private readonly ShellViewModel _shellViewModel;

    private readonly DispatcherTimer _searchDebounceTimer;
    private CancellationTokenSource? _loadCts;

                 
    private bool _isLoading;
    private string? _errorMessage;
    private string? _searchQuery;
    private bool _includeDynamic = true;

             
    private int _totalCount;
    private int _currentSkip;
    private const int PageSize = 50;
    private bool _hasMore;

                            
    private DistributionListItemDto? _selectedItem;
    private DistributionListDetailsDto? _selectedDetails;
    private bool _isLoadingDetails;

                     
    private int _membersCurrentSkip;
    private bool _membersHasMore;
    private bool _isLoadingMembers;

                        
    private string? _newMemberIdentity;

    private string? _newDistributionListDisplayName;
    private string? _newDistributionListAlias;
    private string? _newDistributionListLocalPart;
    private string? _selectedDistributionListDomain;
    private bool _isCreatingDistributionList;

               
    private bool _allowExternalSenders;
    private bool _originalAllowExternalSenders;
    private List<string> _originalAcceptMessagesOnlyFrom = new();
    private List<string> _originalRejectMessagesFrom = new();
    private bool _hasPendingSettingsChanges;
    private bool _isSavingSettings;
    private bool _isInitializingSettings;
    private string? _newAcceptedSender;
    private string? _newRejectedSender;

    public DistributionListViewModel(IWorkerService workerService, NavigationService navigationService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _navigationService = navigationService;
        _shellViewModel = shellViewModel;

                                        
        _searchDebounceTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(300)
        };
        _searchDebounceTimer.Tick += OnSearchDebounceElapsed;

                                       
        _navigationService.SelectedIdentityChanged += OnSelectedIdentityChanged;
        _shellViewModel.PropertyChanged += OnShellViewModelPropertyChanged;

                   
        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
        LoadMoreCommand = new AsyncRelayCommand(LoadMoreAsync, () => CanLoadMore);
        CancelCommand = new RelayCommand(Cancel, () => IsLoading);
        ViewDetailsCommand = new RelayCommand<DistributionListItemDto>(ViewDetails, m => m != null);
        BackCommand = new RelayCommand(GoBack);

        LoadMoreMembersCommand = new AsyncRelayCommand(LoadMoreMembersAsync, () => MembersHasMore && !IsLoadingMembers);
        PreviewDynamicMembersCommand = new AsyncRelayCommand(PreviewDynamicMembersAsync, () => CanPreviewDynamicMembers);
        AddMemberCommand = new AsyncRelayCommand(AddMemberAsync, () => CanAddMember);
        RemoveMemberCommand = new AsyncRelayCommand<GroupMemberDto>(RemoveMemberAsync, m => m != null && CanModifyMembers);
        SaveSettingsCommand = new AsyncRelayCommand(SaveSettingsAsync, () => HasPendingSettingsChanges && !IsSavingSettings);
        DiscardSettingsCommand = new RelayCommand(DiscardSettingsChanges, () => HasPendingSettingsChanges);
        AddAcceptSenderCommand = new RelayCommand(AddAcceptSender, () => CanAddAcceptSender);
        RemoveAcceptSenderCommand = new RelayCommand<string>(RemoveAcceptSender, sender => CanRemoveAcceptSender(sender));
        AddRejectSenderCommand = new RelayCommand(AddRejectSender, () => CanAddRejectSender);
        RemoveRejectSenderCommand = new RelayCommand<string>(RemoveRejectSender, sender => CanRemoveRejectSender(sender));
        CreateDistributionListCommand = new AsyncRelayCommand(CreateDistributionListAsync, () => CanCreateDistributionList);
    }

    #region Properties

    public ObservableCollection<DistributionListItemDto> DistributionLists { get; } = new();
    public ObservableCollection<GroupMemberDto> Members { get; } = new();
    public ObservableCollection<string> AcceptMessagesOnlyFrom { get; } = new();
    public ObservableCollection<string> RejectMessagesFrom { get; } = new();
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
                _searchDebounceTimer.Stop();
                _searchDebounceTimer.Start();
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
                TriggerRefreshFromUi();
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

    public DistributionListItemDto? SelectedItem
    {
        get => _selectedItem;
        set
        {
            if (SetProperty(ref _selectedItem, value))
            {
                OnPropertyChanged(nameof(HasSelection));
                if (value != null)
                {
                    ViewDetails(value);
                }
            }
        }
    }

    public bool HasSelection => SelectedItem != null || SelectedDetails != null;

    public DistributionListDetailsDto? SelectedDetails
    {
        get => _selectedDetails;
        private set
        {
            if (SetProperty(ref _selectedDetails, value))
            {
                OnPropertyChanged(nameof(HasDetails));
                OnPropertyChanged(nameof(IsDynamicGroup));
                OnPropertyChanged(nameof(CanModifyMembers));
                OnPropertyChanged(nameof(CanEditSettings));
                OnPropertyChanged(nameof(CanEditExternalSenders));
                OnPropertyChanged(nameof(CanEditAcceptMessagesOnlyFrom));
                OnPropertyChanged(nameof(CanEditRejectMessagesFrom));
                OnPropertyChanged(nameof(CanEditSenderFilters));
                OnPropertyChanged(nameof(CanAddAcceptSender));
                OnPropertyChanged(nameof(CanAddRejectSender));
                OnPropertyChanged(nameof(CanPreviewDynamicMembers));
                InitializeSettingsFromDetails();
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool HasDetails => SelectedDetails != null;
    public bool IsDynamicGroup => SelectedDetails?.GroupType == "Dynamic";
    public bool CanModifyMembers => HasDetails && !IsDynamicGroup && _shellViewModel.IsFeatureAvailable(f => f.CanAddDistributionGroupMember);
    public bool CanIncludeDynamicFilter => _shellViewModel.IsFeatureAvailable(f => f.CanGetDynamicDistributionGroup);
    public bool CanPreviewDynamicMembers => HasDetails && IsDynamicGroup && !IsLoadingMembers && _shellViewModel.IsFeatureAvailable(f => f.CanGetDynamicDistributionGroup);

    public bool IsLoadingDetails
    {
        get => _isLoadingDetails;
        private set => SetProperty(ref _isLoadingDetails, value);
    }

    public bool MembersHasMore
    {
        get => _membersHasMore;
        private set
        {
            if (SetProperty(ref _membersHasMore, value))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool IsLoadingMembers
    {
        get => _isLoadingMembers;
        private set
        {
            if (SetProperty(ref _isLoadingMembers, value))
            {
                OnPropertyChanged(nameof(CanPreviewDynamicMembers));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool IsSavingSettings
    {
        get => _isSavingSettings;
        private set
        {
            if (SetProperty(ref _isSavingSettings, value))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string? NewMemberIdentity
    {
        get => _newMemberIdentity;
        set
        {
            if (SetProperty(ref _newMemberIdentity, value))
            {
                OnPropertyChanged(nameof(CanAddMember));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool CanAddMember => !string.IsNullOrWhiteSpace(NewMemberIdentity) && CanModifyMembers;
    public bool CanEditSettings => HasDetails && (IsDynamicGroup
        ? _shellViewModel.IsFeatureAvailable(f => f.CanSetDynamicDistributionGroup)
        : _shellViewModel.IsFeatureAvailable(f => f.CanSetDistributionGroup));

    public bool CanEditExternalSenders => CanEditSettings && (IsDynamicGroup
        ? _shellViewModel.IsFeatureAvailable(f => f.CanSetDynamicDistributionGroupRequireSenderAuthentication)
        : _shellViewModel.IsFeatureAvailable(f => f.CanSetDistributionGroupRequireSenderAuthentication));

    public bool CanEditAcceptMessagesOnlyFrom => CanEditSettings && (IsDynamicGroup
        ? _shellViewModel.IsFeatureAvailable(f => f.CanSetDynamicDistributionGroupAcceptMessagesOnlyFrom)
        : _shellViewModel.IsFeatureAvailable(f => f.CanSetDistributionGroupAcceptMessagesOnlyFrom));

    public bool CanEditRejectMessagesFrom => CanEditSettings && (IsDynamicGroup
        ? _shellViewModel.IsFeatureAvailable(f => f.CanSetDynamicDistributionGroupRejectMessagesFrom)
        : _shellViewModel.IsFeatureAvailable(f => f.CanSetDistributionGroupRejectMessagesFrom));

    public bool CanEditSenderFilters => CanEditAcceptMessagesOnlyFrom && CanEditRejectMessagesFrom;

    public string? NewAcceptedSender
    {
        get => _newAcceptedSender;
        set
        {
            if (SetProperty(ref _newAcceptedSender, value))
            {
                OnPropertyChanged(nameof(CanAddAcceptSender));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string? NewRejectedSender
    {
        get => _newRejectedSender;
        set
        {
            if (SetProperty(ref _newRejectedSender, value))
            {
                OnPropertyChanged(nameof(CanAddRejectSender));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool CanAddAcceptSender => CanEditAcceptMessagesOnlyFrom && !string.IsNullOrWhiteSpace(NewAcceptedSender);
    public bool CanAddRejectSender => CanEditRejectMessagesFrom && !string.IsNullOrWhiteSpace(NewRejectedSender);

    public bool AllowExternalSenders
    {
        get => _allowExternalSenders;
        set
        {
            if (SetProperty(ref _allowExternalSenders, value))
            {
                UpdatePendingSettingsChanges();
            }
        }
    }

    public bool HasPendingSettingsChanges
    {
        get => _hasPendingSettingsChanges;
        private set
        {
            if (SetProperty(ref _hasPendingSettingsChanges, value))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }


    public string? NewDistributionListDisplayName
    {
        get => _newDistributionListDisplayName;
        set
        {
            if (SetProperty(ref _newDistributionListDisplayName, value))
            {
                OnPropertyChanged(nameof(CanCreateDistributionList));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string? NewDistributionListAlias
    {
        get => _newDistributionListAlias;
        set
        {
            if (SetProperty(ref _newDistributionListAlias, value))
            {
                OnPropertyChanged(nameof(CanCreateDistributionList));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string? NewDistributionListLocalPart
    {
        get => _newDistributionListLocalPart;
        set
        {
            if (SetProperty(ref _newDistributionListLocalPart, value))
            {
                OnPropertyChanged(nameof(CanCreateDistributionList));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string? SelectedDistributionListDomain
    {
        get => _selectedDistributionListDomain;
        set
        {
            if (SetProperty(ref _selectedDistributionListDomain, value))
            {
                OnPropertyChanged(nameof(CanCreateDistributionList));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool IsCreatingDistributionList
    {
        get => _isCreatingDistributionList;
        private set
        {
            if (SetProperty(ref _isCreatingDistributionList, value))
            {
                OnPropertyChanged(nameof(CanCreateDistributionList));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool CanCreateDistributionList =>
        !IsLoading && !IsCreatingDistributionList && _shellViewModel.IsExchangeConnected &&
        !string.IsNullOrWhiteSpace(NewDistributionListDisplayName) &&
        !string.IsNullOrWhiteSpace(NewDistributionListAlias) &&
        !string.IsNullOrWhiteSpace(NewDistributionListLocalPart) &&
        !string.IsNullOrWhiteSpace(SelectedDistributionListDomain);

    public bool CanRefresh => !IsLoading && _shellViewModel.IsExchangeConnected;
    public bool CanLoadMore => !IsLoading && HasMore && _shellViewModel.IsExchangeConnected;

    public string StatusText => $"{DistributionLists.Count} di {TotalCount} gruppi";
    public string MembersStatusText => SelectedDetails?.Members != null
        ? $"{Members.Count} di {SelectedDetails.Members.TotalCount} membri"
        : $"{Members.Count} membri";

    #endregion

    #region Commands

    public ICommand RefreshCommand { get; }
    public ICommand LoadMoreCommand { get; }
    public ICommand CancelCommand { get; }
    public ICommand ViewDetailsCommand { get; }
    public ICommand BackCommand { get; }
    public ICommand LoadMoreMembersCommand { get; }
    public ICommand PreviewDynamicMembersCommand { get; }
    public ICommand AddMemberCommand { get; }
    public ICommand RemoveMemberCommand { get; }
    public ICommand SaveSettingsCommand { get; }
    public ICommand DiscardSettingsCommand { get; }
    public ICommand AddAcceptSenderCommand { get; }
    public ICommand RemoveAcceptSenderCommand { get; }
    public ICommand AddRejectSenderCommand { get; }
    public ICommand RemoveRejectSenderCommand { get; }
    public ICommand CreateDistributionListCommand { get; }

    #endregion

    #region Methods

    public async Task LoadAsync()
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            DistributionLists.Clear();
            ErrorMessage = "Non connesso a Exchange Online";
            return;
        }

        await LoadAcceptedDomainsAsync(CancellationToken.None);

        if (!CanIncludeDynamicFilter && IncludeDynamic)
        {
            IncludeDynamic = false;
        }

        await RefreshAsync(CancellationToken.None);
    }

    private void OnSelectedIdentityChanged(object? sender, string? identity)
    {
        if (_navigationService.CurrentPage == NavigationPage.DistributionLists)
        {
            if (!string.IsNullOrEmpty(identity))
            {
                _ = LoadDetailsAsync(identity);
            }
            else
            {
                SelectedDetails = null;
                Members.Clear();
            }
        }
    }

    private void OnShellViewModelPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(ShellViewModel.Capabilities)
            or nameof(ShellViewModel.ExchangeState)
            or nameof(ShellViewModel.IsExchangeConnected))
        {
            OnPropertyChanged(nameof(CanRefresh));
            OnPropertyChanged(nameof(CanLoadMore));
            OnPropertyChanged(nameof(CanModifyMembers));
            OnPropertyChanged(nameof(CanEditSettings));
            OnPropertyChanged(nameof(CanCreateDistributionList));
            OnPropertyChanged(nameof(CanEditExternalSenders));
            OnPropertyChanged(nameof(CanEditAcceptMessagesOnlyFrom));
            OnPropertyChanged(nameof(CanEditRejectMessagesFrom));
            OnPropertyChanged(nameof(CanEditSenderFilters));
            OnPropertyChanged(nameof(CanAddAcceptSender));
            OnPropertyChanged(nameof(CanAddRejectSender));
            OnPropertyChanged(nameof(CanIncludeDynamicFilter));
            OnPropertyChanged(nameof(CanPreviewDynamicMembers));

            if (!CanIncludeDynamicFilter && IncludeDynamic)
            {
                IncludeDynamic = false;
            }

            UpdatePendingSettingsChanges();
            CommandManager.InvalidateRequerySuggested();
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
            _shellViewModel.AddLog(LogLevel.Error, $"Distribution list refresh failed: {ex.Message}");
        }
    }

    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        _loadCts?.Cancel();
        _loadCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

        IsLoading = true;
        ErrorMessage = null;
        _currentSkip = 0;
        DistributionLists.Clear();

        try
        {
            var request = new GetDistributionListsRequest
            {
                SearchQuery = string.IsNullOrWhiteSpace(_searchQuery) ? null : _searchQuery.Trim(),
                IncludeDynamic = _includeDynamic && CanIncludeDynamicFilter,
                PageSize = PageSize,
                Skip = 0
            };

            var result = await _workerService.GetDistributionListsAsync(
                request,
                eventHandler: null,
                cancellationToken: _loadCts.Token);

            if (result.IsSuccess && result.Value != null)
            {
                foreach (var item in result.Value.DistributionLists)
                {
                    DistributionLists.Add(item);
                }

                TotalCount = result.Value.TotalCount;
                HasMore = result.Value.HasMore;
                _currentSkip = result.Value.DistributionLists.Count;

                OnPropertyChanged(nameof(StatusText));
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? "Impossibile caricare le liste di distribuzione";
                _shellViewModel.AddLog(LogLevel.Error, $"Distribution list load failed: {ErrorMessage}");
            }
        }
        catch (OperationCanceledException)
        {
                     
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Distribution list exception: {ex.Message}");
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
        _loadCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

        IsLoading = true;
        ErrorMessage = null;

        try
        {
            var request = new GetDistributionListsRequest
            {
                SearchQuery = string.IsNullOrWhiteSpace(_searchQuery) ? null : _searchQuery.Trim(),
                IncludeDynamic = _includeDynamic && CanIncludeDynamicFilter,
                PageSize = PageSize,
                Skip = _currentSkip
            };

            var result = await _workerService.GetDistributionListsAsync(
                request,
                eventHandler: null,
                cancellationToken: _loadCts.Token);

            if (result.IsSuccess && result.Value != null)
            {
                foreach (var item in result.Value.DistributionLists)
                {
                    DistributionLists.Add(item);
                }

                HasMore = result.Value.HasMore;
                _currentSkip += result.Value.DistributionLists.Count;

                OnPropertyChanged(nameof(StatusText));
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? "Failed to load more";
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
        _ = SafeRefreshAsync();
    }

    private void ViewDetails(DistributionListItemDto? item)
    {
        if (item == null) return;
        _navigationService.NavigateToDetails(NavigationPage.DistributionLists, item.Identity, item);
    }

    private void GoBack()
    {
        SelectedDetails = null;
        Members.Clear();
        _navigationService.ClearSelection();
    }

    private async Task LoadDetailsAsync(string identity)
    {
        IsLoadingDetails = true;
        ErrorMessage = null;
        Members.Clear();
        _membersCurrentSkip = 0;

        try
        {
            var request = new GetDistributionListDetailsRequest
            {
                Identity = identity,
                IncludeMembers = true,
                MembersPageSize = PageSize
            };

            var result = await _workerService.GetDistributionListDetailsAsync(
                request,
                eventHandler: null,
                cancellationToken: CancellationToken.None);

            if (result.IsSuccess && result.Value != null)
            {
                SelectedDetails = result.Value;

                                   
                if (result.Value.Members != null)
                {
                    foreach (var member in result.Value.Members.Members)
                    {
                        Members.Add(member);
                    }
                    MembersHasMore = result.Value.Members.HasMore;
                    _membersCurrentSkip = result.Value.Members.Members.Count;
                }

                OnPropertyChanged(nameof(MembersStatusText));
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? "Impossibile caricare i dettagli";
                _shellViewModel.AddLog(LogLevel.Error, $"Distribution list details load failed: {ErrorMessage}");
            }
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Distribution list details error: {ex.Message}");
        }
        finally
        {
            IsLoadingDetails = false;
        }
    }


    private string BuildDistributionListPrimarySmtpAddress()
    {
        var localPart = NewDistributionListLocalPart?.Trim() ?? string.Empty;
        var domain = SelectedDistributionListDomain?.Trim() ?? string.Empty;
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

            AvailableMailDomains.Clear();
            foreach (var domain in domains)
            {
                AvailableMailDomains.Add(domain!);
            }

            if (string.IsNullOrWhiteSpace(SelectedDistributionListDomain))
            {
                SelectedDistributionListDomain = result.Value.Domains.FirstOrDefault(d => d.Default)?.DomainName
                    ?? AvailableMailDomains.FirstOrDefault();
            }
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Warning, $"Impossibile caricare i domini accettati: {ex.Message}");
        }
    }

    private async Task CreateDistributionListAsync(CancellationToken cancellationToken)
    {
        if (!CanCreateDistributionList)
        {
            return;
        }

        IsCreatingDistributionList = true;
        ErrorMessage = null;

        try
        {
            var request = new CreateDistributionListRequest
            {
                DisplayName = NewDistributionListDisplayName!.Trim(),
                Alias = NewDistributionListAlias!.Trim(),
                PrimarySmtpAddress = BuildDistributionListPrimarySmtpAddress()
            };

            _shellViewModel.AddLog(LogLevel.Information, $"Creazione lista {request.PrimarySmtpAddress}...");

            var result = await _workerService.CreateDistributionListAsync(request, cancellationToken: cancellationToken);
            if (result.IsSuccess)
            {
                _shellViewModel.AddLog(LogLevel.Information, "Lista di distribuzione creata con successo");
                NewDistributionListDisplayName = string.Empty;
                NewDistributionListAlias = string.Empty;
                NewDistributionListLocalPart = string.Empty;
                await RefreshAsync(cancellationToken);
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? "Impossibile creare la lista di distribuzione";
                _shellViewModel.AddLog(LogLevel.Error, $"Creazione lista fallita: {ErrorMessage}");
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Errore creazione lista: {ex.Message}");
        }
        finally
        {
            IsCreatingDistributionList = false;
        }
    }

    private async Task LoadMoreMembersAsync(CancellationToken cancellationToken)
    {
        if (!MembersHasMore || IsLoadingMembers || SelectedDetails == null) return;

        IsLoadingMembers = true;

        try
        {
            var request = new GetGroupMembersRequest
            {
                Identity = SelectedDetails.Identity,
                GroupType = IsDynamicGroup ? "DynamicDistributionGroup" : "DistributionGroup",
                PageSize = PageSize,
                Skip = _membersCurrentSkip
            };

            var result = await _workerService.GetGroupMembersAsync(
                request,
                eventHandler: null,
                cancellationToken: cancellationToken);

            if (result.IsSuccess && result.Value != null)
            {
                foreach (var member in result.Value.Members)
                {
                    Members.Add(member);
                }
                MembersHasMore = result.Value.HasMore;
                _membersCurrentSkip += result.Value.Members.Count;

                OnPropertyChanged(nameof(MembersStatusText));
            }
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Load more members error: {ex.Message}");
        }
        finally
        {
            IsLoadingMembers = false;
        }
    }

    private async Task PreviewDynamicMembersAsync(CancellationToken cancellationToken)
    {
        if (SelectedDetails == null || !IsDynamicGroup) return;

        IsLoadingMembers = true;
        Members.Clear();

        try
        {
            var request = new PreviewDynamicGroupMembersRequest
            {
                Identity = SelectedDetails.Identity,
                MaxResults = 100
            };

            _shellViewModel.AddLog(LogLevel.Warning, "Previewing dynamic group members (this may take a while)...");

            var result = await _workerService.PreviewDynamicGroupMembersAsync(
                request,
                eventHandler: null,
                cancellationToken: cancellationToken);

            if (result.IsSuccess && result.Value != null)
            {
                foreach (var member in result.Value.Members)
                {
                    Members.Add(member);
                }

                if (result.Value.Warning != null)
                {
                    _shellViewModel.AddLog(LogLevel.Warning, result.Value.Warning);
                }

                _shellViewModel.AddLog(LogLevel.Information, $"Preview complete: {result.Value.Members.Count} of {result.Value.TotalCount} members");
            }
            else if (!result.WasCancelled)
            {
                _shellViewModel.AddLog(LogLevel.Error, $"Preview failed: {result.Error?.Message}");
            }
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Preview error: {ex.Message}");
        }
        finally
        {
            IsLoadingMembers = false;
        }
    }

    private async Task AddMemberAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(NewMemberIdentity) || SelectedDetails == null || IsDynamicGroup) return;

        try
        {
            var request = new ModifyGroupMemberRequest
            {
                Identity = SelectedDetails.Identity,
                Member = NewMemberIdentity.Trim(),
                Action = GroupMemberAction.Add,
                GroupType = "DistributionGroup"
            };

            _shellViewModel.AddLog(LogLevel.Information, $"Adding {NewMemberIdentity} to {SelectedDetails.DisplayName}...");

            var result = await _workerService.ModifyGroupMemberAsync(
                request,
                eventHandler: null,
                cancellationToken: cancellationToken);

            if (result.IsSuccess)
            {
                _shellViewModel.AddLog(LogLevel.Information, "Member added successfully");
                NewMemberIdentity = string.Empty;

                                  
                Members.Clear();
                _membersCurrentSkip = 0;
                await LoadDetailsAsync(SelectedDetails.Identity);
            }
            else
            {
                _shellViewModel.AddLog(LogLevel.Error, $"Add member failed: {result.Error?.Message}");
            }
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Add member error: {ex.Message}");
        }
    }

    private async Task RemoveMemberAsync(GroupMemberDto? member)
    {
        if (member == null || SelectedDetails == null || IsDynamicGroup) return;

        try
        {
            var request = new ModifyGroupMemberRequest
            {
                Identity = SelectedDetails.Identity,
                Member = member.Identity,
                Action = GroupMemberAction.Remove,
                GroupType = "DistributionGroup"
            };

            _shellViewModel.AddLog(LogLevel.Information, $"Removing {member.Name} from {SelectedDetails.DisplayName}...");

            var result = await _workerService.ModifyGroupMemberAsync(
                request,
                eventHandler: null,
                cancellationToken: CancellationToken.None);

            if (result.IsSuccess)
            {
                _shellViewModel.AddLog(LogLevel.Information, "Member removed successfully");
                Members.Remove(member);
                OnPropertyChanged(nameof(MembersStatusText));
            }
            else
            {
                _shellViewModel.AddLog(LogLevel.Error, $"Remove member failed: {result.Error?.Message}");
            }
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Remove member error: {ex.Message}");
        }
    }

    public void Cancel()
    {
        _loadCts?.Cancel();
    }

    private void LoadSenderFilters(DistributionListDetailsDto details)
    {
        AcceptMessagesOnlyFrom.Clear();
        foreach (var sender in details.AcceptMessagesOnlyFrom)
        {
            AcceptMessagesOnlyFrom.Add(sender);
        }

        RejectMessagesFrom.Clear();
        foreach (var sender in details.RejectMessagesFrom)
        {
            RejectMessagesFrom.Add(sender);
        }

        _originalAcceptMessagesOnlyFrom = NormalizeSenderList(details.AcceptMessagesOnlyFrom).ToList();
        _originalRejectMessagesFrom = NormalizeSenderList(details.RejectMessagesFrom).ToList();
    }

    private void AddAcceptSender()
    {
        if (!CanAddAcceptSender)
        {
            return;
        }

        var normalized = NewAcceptedSender!.Trim();
        if (!AcceptMessagesOnlyFrom.Any(s => string.Equals(s, normalized, StringComparison.OrdinalIgnoreCase)))
        {
            AcceptMessagesOnlyFrom.Add(normalized);
            UpdatePendingSettingsChanges();
        }

        NewAcceptedSender = string.Empty;
    }

    private void RemoveAcceptSender(string? sender)
    {
        if (!CanRemoveAcceptSender(sender))
        {
            return;
        }

        AcceptMessagesOnlyFrom.Remove(sender!);
        UpdatePendingSettingsChanges();
    }

    private void AddRejectSender()
    {
        if (!CanAddRejectSender)
        {
            return;
        }

        var normalized = NewRejectedSender!.Trim();
        if (!RejectMessagesFrom.Any(s => string.Equals(s, normalized, StringComparison.OrdinalIgnoreCase)))
        {
            RejectMessagesFrom.Add(normalized);
            UpdatePendingSettingsChanges();
        }

        NewRejectedSender = string.Empty;
    }

    private void RemoveRejectSender(string? sender)
    {
        if (!CanRemoveRejectSender(sender))
        {
            return;
        }

        RejectMessagesFrom.Remove(sender!);
        UpdatePendingSettingsChanges();
    }

    private bool CanRemoveAcceptSender(string? sender) => CanEditAcceptMessagesOnlyFrom && !string.IsNullOrWhiteSpace(sender);

    private bool CanRemoveRejectSender(string? sender) => CanEditRejectMessagesFrom && !string.IsNullOrWhiteSpace(sender);

    private static IEnumerable<string> NormalizeSenderList(IEnumerable<string> list)
    {
        return list
            .Select(value => value?.Trim())
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)!;
    }

    private static bool SenderListEquals(IEnumerable<string> current, IEnumerable<string> original)
    {
        var normalizedCurrent = NormalizeSenderList(current).ToList();
        var normalizedOriginal = NormalizeSenderList(original).ToList();
        return normalizedCurrent.SequenceEqual(normalizedOriginal, StringComparer.OrdinalIgnoreCase);
    }

    private void InitializeSettingsFromDetails()
    {
        _isInitializingSettings = true;

        if (SelectedDetails != null)
        {
            AllowExternalSenders = !SelectedDetails.RequireSenderAuthenticationEnabled;
            _originalAllowExternalSenders = AllowExternalSenders;
            LoadSenderFilters(SelectedDetails);
            NewAcceptedSender = string.Empty;
            NewRejectedSender = string.Empty;
        }
        else
        {
            AllowExternalSenders = false;
            _originalAllowExternalSenders = false;
            AcceptMessagesOnlyFrom.Clear();
            RejectMessagesFrom.Clear();
            _originalAcceptMessagesOnlyFrom = new List<string>();
            _originalRejectMessagesFrom = new List<string>();
            NewAcceptedSender = string.Empty;
            NewRejectedSender = string.Empty;
        }

        _isInitializingSettings = false;
        HasPendingSettingsChanges = false;
    }

    private void UpdatePendingSettingsChanges()
    {
        if (_isInitializingSettings)
        {
            return;
        }

        var externalChanged = CanEditExternalSenders && AllowExternalSenders != _originalAllowExternalSenders;
        var acceptChanged = CanEditAcceptMessagesOnlyFrom && !SenderListEquals(AcceptMessagesOnlyFrom, _originalAcceptMessagesOnlyFrom);
        var rejectChanged = CanEditRejectMessagesFrom && !SenderListEquals(RejectMessagesFrom, _originalRejectMessagesFrom);

        HasPendingSettingsChanges = externalChanged || acceptChanged || rejectChanged;
    }

    private async Task SaveSettingsAsync(CancellationToken cancellationToken)
    {
        if (SelectedDetails == null || !CanEditSettings)
        {
            return;
        }

        IsSavingSettings = true;
        ErrorMessage = null;

        try
        {
            var acceptChanged = !SenderListEquals(AcceptMessagesOnlyFrom, _originalAcceptMessagesOnlyFrom);
            var rejectChanged = !SenderListEquals(RejectMessagesFrom, _originalRejectMessagesFrom);

            var request = new SetDistributionListSettingsRequest
            {
                Identity = SelectedDetails.Identity,
                GroupType = IsDynamicGroup ? "DynamicDistributionGroup" : "DistributionGroup",
                RequireSenderAuthenticationEnabled = CanEditExternalSenders ? !AllowExternalSenders : null,
                AcceptMessagesOnlyFrom = (CanEditAcceptMessagesOnlyFrom && acceptChanged) ? NormalizeSenderList(AcceptMessagesOnlyFrom).ToList() : null,
                RejectMessagesFrom = (CanEditRejectMessagesFrom && rejectChanged) ? NormalizeSenderList(RejectMessagesFrom).ToList() : null
            };

            _shellViewModel.AddLog(LogLevel.Information, $"Aggiornamento impostazioni lista {SelectedDetails.DisplayName}...");

            var result = await _workerService.SetDistributionListSettingsAsync(request, cancellationToken: cancellationToken);

            if (result.IsSuccess)
            {
                if (CanEditExternalSenders)
                {
                    SelectedDetails.RequireSenderAuthenticationEnabled = !AllowExternalSenders;
                    _originalAllowExternalSenders = AllowExternalSenders;
                }

                if (CanEditAcceptMessagesOnlyFrom)
                {
                    SelectedDetails.AcceptMessagesOnlyFrom = NormalizeSenderList(AcceptMessagesOnlyFrom).ToList();
                    _originalAcceptMessagesOnlyFrom = NormalizeSenderList(AcceptMessagesOnlyFrom).ToList();
                }

                if (CanEditRejectMessagesFrom)
                {
                    SelectedDetails.RejectMessagesFrom = NormalizeSenderList(RejectMessagesFrom).ToList();
                    _originalRejectMessagesFrom = NormalizeSenderList(RejectMessagesFrom).ToList();
                }
                HasPendingSettingsChanges = false;
                _shellViewModel.AddLog(LogLevel.Information, "Impostazioni lista aggiornate");
            }
            else
            {
                ErrorMessage = result.Error?.Message ?? "Impossibile aggiornare le impostazioni della lista";
                _shellViewModel.AddLog(LogLevel.Error, $"Aggiornamento lista fallito: {ErrorMessage}");
            }
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Errore aggiornamento lista: {ex.Message}");
        }
        finally
        {
            IsSavingSettings = false;
        }
    }

    private void DiscardSettingsChanges()
    {
        if (_isInitializingSettings)
        {
            return;
        }

        AllowExternalSenders = _originalAllowExternalSenders;
        AcceptMessagesOnlyFrom.Clear();
        foreach (var sender in _originalAcceptMessagesOnlyFrom)
        {
            AcceptMessagesOnlyFrom.Add(sender);
        }

        RejectMessagesFrom.Clear();
        foreach (var sender in _originalRejectMessagesFrom)
        {
            RejectMessagesFrom.Add(sender);
        }

        HasPendingSettingsChanges = false;
        _shellViewModel.AddLog(LogLevel.Information, "Modifiche lista annullate");
    }

    #endregion
}
