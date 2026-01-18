using System.Collections.ObjectModel;
using System.Windows.Input;
using System.Windows.Threading;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

/// <summary>
/// Distribution list view model with paging and search.
/// </summary>
public class DistributionListViewModel : ViewModelBase
{
    private readonly IWorkerService _workerService;
    private readonly NavigationService _navigationService;
    private readonly ShellViewModel _shellViewModel;

    private readonly DispatcherTimer _searchDebounceTimer;
    private CancellationTokenSource? _loadCts;

    // List state
    private bool _isLoading;
    private string? _errorMessage;
    private string? _searchQuery;
    private bool _includeDynamic = true;

    // Paging
    private int _totalCount;
    private int _currentSkip;
    private const int PageSize = 50;
    private bool _hasMore;

    // Selection and details
    private DistributionListItemDto? _selectedItem;
    private DistributionListDetailsDto? _selectedDetails;
    private bool _isLoadingDetails;

    // Members paging
    private int _membersCurrentSkip;
    private bool _membersHasMore;
    private bool _isLoadingMembers;

    // Add/remove member
    private string? _newMemberIdentity;

    public DistributionListViewModel(IWorkerService workerService, NavigationService navigationService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _navigationService = navigationService;
        _shellViewModel = shellViewModel;

        // Search debounce timer (300ms)
        _searchDebounceTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(300)
        };
        _searchDebounceTimer.Tick += OnSearchDebounceElapsed;

        // Listen for selection changes
        _navigationService.SelectedIdentityChanged += OnSelectedIdentityChanged;

        // Commands
        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
        LoadMoreCommand = new AsyncRelayCommand(LoadMoreAsync, () => CanLoadMore);
        CancelCommand = new RelayCommand(Cancel, () => IsLoading);
        ViewDetailsCommand = new RelayCommand<DistributionListItemDto>(ViewDetails, m => m != null);
        BackCommand = new RelayCommand(GoBack);

        LoadMoreMembersCommand = new AsyncRelayCommand(LoadMoreMembersAsync, () => MembersHasMore && !IsLoadingMembers);
        PreviewDynamicMembersCommand = new AsyncRelayCommand(PreviewDynamicMembersAsync, () => SelectedDetails != null && SelectedDetails.GroupType == "Dynamic");
        AddMemberCommand = new AsyncRelayCommand(AddMemberAsync, () => CanAddMember);
        RemoveMemberCommand = new AsyncRelayCommand<GroupMemberDto>(RemoveMemberAsync, m => m != null && CanModifyMembers);
    }

    #region Properties

    public ObservableCollection<DistributionListItemDto> DistributionLists { get; } = new();
    public ObservableCollection<GroupMemberDto> Members { get; } = new();

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
            if (SetProperty(ref _includeDynamic, value))
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
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool HasDetails => SelectedDetails != null;
    public bool IsDynamicGroup => SelectedDetails?.GroupType == "Dynamic";
    public bool CanModifyMembers => HasDetails && !IsDynamicGroup && _shellViewModel.IsFeatureAvailable(f => f.CanAddDistributionGroupMember);

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

    public bool CanRefresh => !IsLoading && _shellViewModel.IsExchangeConnected;
    public bool CanLoadMore => !IsLoading && HasMore && _shellViewModel.IsExchangeConnected;

    public string StatusText => $"{DistributionLists.Count} of {TotalCount} groups";
    public string MembersStatusText => SelectedDetails?.Members != null
        ? $"{Members.Count} of {SelectedDetails.Members.TotalCount} members"
        : $"{Members.Count} members";

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

    #endregion

    #region Methods

    public async Task LoadAsync()
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            DistributionLists.Clear();
            ErrorMessage = "Not connected to Exchange Online";
            return;
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
            Console.WriteLine("[DistributionListViewModel] Requesting distribution lists...");

            var request = new GetDistributionListsRequest
            {
                SearchQuery = _searchQuery,
                IncludeDynamic = _includeDynamic,
                PageSize = PageSize,
                Skip = 0
            };

            var result = await _workerService.GetDistributionListsAsync(
                request,
                eventHandler: null,
                cancellationToken: _loadCts.Token);

            Console.WriteLine($"[DistributionListViewModel] Response received - IsSuccess: {result.IsSuccess}, HasValue: {result.Value != null}");

            if (result.IsSuccess && result.Value != null)
            {
                Console.WriteLine($"[DistributionListViewModel] Groups loaded - Count: {result.Value.DistributionLists.Count}, TotalCount: {result.Value.TotalCount}, HasMore: {result.Value.HasMore}");

                foreach (var item in result.Value.DistributionLists)
                {
                    DistributionLists.Add(item);
                }

                TotalCount = result.Value.TotalCount;
                HasMore = result.Value.HasMore;
                _currentSkip = result.Value.DistributionLists.Count;

                OnPropertyChanged(nameof(StatusText));
                Console.WriteLine($"[DistributionListViewModel] Successfully loaded {DistributionLists.Count} groups. ErrorMessage = '{ErrorMessage}', HasError = {HasError}");
            }
            else if (result.WasCancelled)
            {
                Console.WriteLine("[DistributionListViewModel] Request was cancelled");
            }
            else
            {
                var errorDetails = result.Error != null
                    ? $"{result.Error.Code}: {result.Error.Message}"
                    : "Failed to load distribution lists (no error details)";
                ErrorMessage = errorDetails;
                _shellViewModel.AddLog(LogLevel.Error, $"Distribution list load failed: {errorDetails}");
                Console.WriteLine($"[DistributionListViewModel] Error: {errorDetails}");
            }
        }
        catch (OperationCanceledException)
        {
            // Ignore
        }
        catch (Exception ex)
        {
            ErrorMessage = $"Exception: {ex.GetType().Name} - {ex.Message}";
            _shellViewModel.AddLog(LogLevel.Error, $"Distribution list exception: {ex.GetType().Name} - {ex.Message}");
            Console.WriteLine($"[DistributionListViewModel] Exception: {ex}");
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
                SearchQuery = _searchQuery,
                IncludeDynamic = _includeDynamic,
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
            // Ignore
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

    private void ViewDetails(DistributionListItemDto? item)
    {
        if (item == null) return;
        SelectedItem = item;
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

                // Populate members
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
                ErrorMessage = result.Error?.Message ?? "Failed to load details";
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

                // Refresh members
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

    #endregion
}
