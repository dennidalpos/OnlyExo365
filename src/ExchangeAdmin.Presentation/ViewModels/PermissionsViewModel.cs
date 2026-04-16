using System.Collections.ObjectModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Paging;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;

namespace ExchangeAdmin.Presentation.ViewModels;

public class PermissionsViewModel : ViewModelBase
{
    private const int PageSize = PagingDefaults.DefaultPageSize;

    private readonly IWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;
    private readonly DebounceHelper _searchDebounce = new();

    private bool _isLoading;
    private bool _isLoadingDetails;
    private bool _isSaving;
    private string? _errorMessage;
    private string? _searchQuery;
    private int _totalCount;
    private int _currentSkip;
    private bool _hasMore;
    private RoleGroupListItemDto? _selectedRoleGroup;
    private RoleGroupDetailsDto? _selectedRoleGroupDetails;
    private string? _newRoleGroupName;
    private string? _newRoleGroupDescription;
    private string? _newRoleGroupRoles;
    private string? _newRoleGroupMembers;
    private string? _copyFromRoleGroupIdentity;
    private string? _memberToAdd;
    private RoleGroupMemberDto? _selectedMember;

    public PermissionsViewModel(IWorkerService workerService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;

        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
        LoadMoreCommand = new AsyncRelayCommand(LoadMoreAsync, () => CanLoadMore);
        NewRoleGroupCommand = new RelayCommand(ResetEditor);
        SaveRoleGroupCommand = new AsyncRelayCommand(SaveRoleGroupAsync, () => CanSaveRoleGroup);
        AddMemberCommand = new AsyncRelayCommand(AddMemberAsync, () => CanAddMember);
        RemoveSelectedMemberCommand = new AsyncRelayCommand(RemoveSelectedMemberAsync, () => CanRemoveSelectedMember);
    }

    public ObservableCollection<RoleGroupListItemDto> RoleGroups { get; } = new();
    public CollectionLoadProgressViewModel LoadProgress { get; } = new();

    public bool IsLoading
    {
        get => _isLoading;
        private set
        {
            if (SetProperty(ref _isLoading, value))
            {
                RaiseCanExecuteChanged();
            }
        }
    }

    public bool IsLoadingDetails
    {
        get => _isLoadingDetails;
        private set
        {
            if (SetProperty(ref _isLoadingDetails, value))
            {
                RaiseCanExecuteChanged();
            }
        }
    }

    public bool IsSaving
    {
        get => _isSaving;
        private set
        {
            if (SetProperty(ref _isSaving, value))
            {
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

    public RoleGroupListItemDto? SelectedRoleGroup
    {
        get => _selectedRoleGroup;
        set
        {
            if (SetProperty(ref _selectedRoleGroup, value))
            {
                OnPropertyChanged(nameof(SelectedRoleGroupSummary));
                RaiseCanExecuteChanged();
                _ = SafeLoadDetailsAsync(value);
            }
        }
    }

    public RoleGroupDetailsDto? SelectedRoleGroupDetails
    {
        get => _selectedRoleGroupDetails;
        private set
        {
            if (SetProperty(ref _selectedRoleGroupDetails, value))
            {
                SelectedMember = null;
                OnPropertyChanged(nameof(RolesText));
                OnPropertyChanged(nameof(MembersCountText));
                OnPropertyChanged(nameof(ScopesText));
                RaiseCanExecuteChanged();
            }
        }
    }

    public string? NewRoleGroupName
    {
        get => _newRoleGroupName;
        set
        {
            if (SetProperty(ref _newRoleGroupName, value))
            {
                RaiseCanExecuteChanged();
            }
        }
    }

    public string? NewRoleGroupDescription
    {
        get => _newRoleGroupDescription;
        set => SetProperty(ref _newRoleGroupDescription, value);
    }

    public string? NewRoleGroupRoles
    {
        get => _newRoleGroupRoles;
        set
        {
            if (SetProperty(ref _newRoleGroupRoles, value))
            {
                RaiseCanExecuteChanged();
            }
        }
    }

    public string? NewRoleGroupMembers
    {
        get => _newRoleGroupMembers;
        set => SetProperty(ref _newRoleGroupMembers, value);
    }

    public string? CopyFromRoleGroupIdentity
    {
        get => _copyFromRoleGroupIdentity;
        set
        {
            if (SetProperty(ref _copyFromRoleGroupIdentity, value))
            {
                OnPropertyChanged(nameof(IsCopyMode));
                RaiseCanExecuteChanged();
            }
        }
    }

    public string? MemberToAdd
    {
        get => _memberToAdd;
        set
        {
            if (SetProperty(ref _memberToAdd, value))
            {
                RaiseCanExecuteChanged();
            }
        }
    }

    public RoleGroupMemberDto? SelectedMember
    {
        get => _selectedMember;
        set
        {
            if (SetProperty(ref _selectedMember, value))
            {
                RaiseCanExecuteChanged();
            }
        }
    }

    public string StatusText => $"{RoleGroups.Count} of {TotalCount} role groups";
    public string SelectedRoleGroupSummary => SelectedRoleGroup == null
        ? "Select a role group to review RBAC roles, membership, and scope."
        : $"{SelectedRoleGroup.DisplayName} - {SelectedRoleGroup.MemberCount} members, {SelectedRoleGroup.RoleCount} roles";

    public string RolesText => SelectedRoleGroupDetails == null || SelectedRoleGroupDetails.Roles.Count == 0
        ? "(none)"
        : string.Join(", ", SelectedRoleGroupDetails.Roles);

    public string MembersCountText => SelectedRoleGroupDetails == null
        ? "-"
        : $"{SelectedRoleGroupDetails.Members.Count} members";

    public string ScopesText => SelectedRoleGroupDetails == null
        ? "-"
        : $"Read: {SelectedRoleGroupDetails.RecipientReadScope ?? "-"} | Write: {SelectedRoleGroupDetails.RecipientWriteScope ?? "-"} | Custom recipient: {SelectedRoleGroupDetails.CustomRecipientWriteScope ?? "-"} | Custom config: {SelectedRoleGroupDetails.CustomConfigWriteScope ?? "-"}";

    public bool IsCopyMode => !string.IsNullOrWhiteSpace(CopyFromRoleGroupIdentity);
    public bool CanRefresh => !IsLoading && !IsLoadingDetails && !IsSaving && _shellViewModel.IsExchangeConnected;
    public bool CanLoadMore => !IsLoading && !IsLoadingDetails && !IsSaving && HasMore && _shellViewModel.IsExchangeConnected;
    public bool CanSaveRoleGroup => !IsLoading && !IsLoadingDetails && !IsSaving && _shellViewModel.IsExchangeConnected &&
                                    !string.IsNullOrWhiteSpace(NewRoleGroupName) &&
                                    (IsCopyMode || !string.IsNullOrWhiteSpace(NewRoleGroupRoles));
    public bool CanAddMember => !IsLoading && !IsLoadingDetails && !IsSaving && _shellViewModel.IsExchangeConnected &&
                                SelectedRoleGroup != null &&
                                !string.IsNullOrWhiteSpace(MemberToAdd);
    public bool CanRemoveSelectedMember => !IsLoading && !IsLoadingDetails && !IsSaving && _shellViewModel.IsExchangeConnected &&
                                           SelectedRoleGroup != null &&
                                           !string.IsNullOrWhiteSpace(SelectedMember?.Identity);

    public ICommand RefreshCommand { get; }
    public ICommand LoadMoreCommand { get; }
    public ICommand NewRoleGroupCommand { get; }
    public ICommand SaveRoleGroupCommand { get; }
    public ICommand AddMemberCommand { get; }
    public ICommand RemoveSelectedMemberCommand { get; }

    public async Task LoadAsync()
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            ResetDisconnectedState();
            return;
        }

        if (RoleGroups.Count == 0)
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
            _shellViewModel.AddLog(LogLevel.Error, $"Permissions refresh failed: {ex.Message}", "Permissions");
        }
    }

    private async Task SafeLoadDetailsAsync(RoleGroupListItemDto? group)
    {
        if (group == null || !_shellViewModel.IsExchangeConnected)
        {
            SelectedRoleGroupDetails = null;
            return;
        }

        try
        {
            await LoadDetailsAsync(group.Identity, CancellationToken.None);
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Permissions details failed: {ex.Message}", "Permissions");
        }
    }

    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            ResetDisconnectedState();
            return;
        }

        IsLoading = true;
        LoadProgress.Start("Loading role groups...", "role groups");
        ErrorMessage = null;
        var refreshPageSize = GetRefreshPageSize(RoleGroups.Count);
        _currentSkip = 0;
        var previousIdentity = SelectedRoleGroup?.Identity;

        try
        {
            var result = await _workerService.GetRoleGroupsAsync(
                new GetRoleGroupsRequest
                {
                    SearchQuery = string.IsNullOrWhiteSpace(SearchQuery) ? null : SearchQuery.Trim(),
                    PageSize = refreshPageSize,
                    Skip = 0,
                    SortBy = "DisplayName"
                },
                eventHandler: HandleLoadProgressEvent,
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to load role groups";
                return;
            }

            RoleGroups.ReplaceAll(result.Value.RoleGroups);
            TotalCount = result.Value.TotalCount;
            HasMore = result.Value.HasMore;
            _currentSkip = RoleGroups.Count;
            SelectedRoleGroup = RestoreSelection(previousIdentity);
            if (SelectedRoleGroup == null)
            {
                SelectedRoleGroupDetails = null;
            }
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
        if (!HasMore)
        {
            return;
        }

        IsLoading = true;
        LoadProgress.Start("Loading role groups...", "role groups");
        ErrorMessage = null;

        try
        {
            var result = await _workerService.GetRoleGroupsAsync(
                new GetRoleGroupsRequest
                {
                    SearchQuery = string.IsNullOrWhiteSpace(SearchQuery) ? null : SearchQuery.Trim(),
                    PageSize = PageSize,
                    Skip = _currentSkip,
                    SortBy = "DisplayName"
                },
                eventHandler: HandleLoadProgressEvent,
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to load role groups";
                return;
            }

            foreach (var item in result.Value.RoleGroups)
            {
                RoleGroups.Add(item);
            }

            TotalCount = result.Value.TotalCount;
            HasMore = result.Value.HasMore;
            _currentSkip = RoleGroups.Count;
            OnPropertyChanged(nameof(StatusText));
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

    private async Task LoadDetailsAsync(string identity, CancellationToken cancellationToken)
    {
        IsLoadingDetails = true;
        ErrorMessage = null;

        try
        {
            var result = await _workerService.GetRoleGroupDetailsAsync(
                new GetRoleGroupDetailsRequest { Identity = identity },
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to load role group details.";
                return;
            }

            if (SelectedRoleGroup == null || !string.Equals(SelectedRoleGroup.Identity, identity, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            SelectedRoleGroupDetails = result.Value;
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsLoadingDetails = false;
        }
    }

    private async Task SaveRoleGroupAsync(CancellationToken cancellationToken)
    {
        if (!ConfirmMutation(
                "Saving role group",
                NewRoleGroupName?.Trim() ?? "New role group",
                "Create or update membership and roles for the selected role group.",
                "Confirm role group save"))
        {
            return;
        }

        IsSaving = true;
        ErrorMessage = null;

        try
        {
            var result = await _workerService.UpsertRoleGroupAsync(
                new UpsertRoleGroupRequest
                {
                    Name = NewRoleGroupName!.Trim(),
                    Description = string.IsNullOrWhiteSpace(NewRoleGroupDescription) ? null : NewRoleGroupDescription.Trim(),
                    Roles = ParseCsv(NewRoleGroupRoles),
                    Members = ParseCsv(NewRoleGroupMembers),
                    CopyFromRoleGroup = string.IsNullOrWhiteSpace(CopyFromRoleGroupIdentity) ? null : CopyFromRoleGroupIdentity
                },
                cancellationToken: cancellationToken);

            if (!result.IsSuccess)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to save the role group.";
                return;
            }

            _shellViewModel.AddLog(LogLevel.Information, $"Role group saved: {NewRoleGroupName}", "Permissions");
            var createdName = NewRoleGroupName;
            ResetEditor();
            await RefreshAsync(cancellationToken);

            var match = RoleGroups.FirstOrDefault(group =>
                string.Equals(group.Name, createdName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(group.DisplayName, createdName, StringComparison.OrdinalIgnoreCase));
            if (match != null)
            {
                SelectedRoleGroup = match;
            }
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsSaving = false;
        }
    }

    private async Task AddMemberAsync(CancellationToken cancellationToken)
    {
        if (SelectedRoleGroup == null || string.IsNullOrWhiteSpace(MemberToAdd))
        {
            return;
        }

        await UpdateMemberAsync(MemberToAdd.Trim(), RoleGroupMemberAction.Add, cancellationToken);
        MemberToAdd = null;
    }

    private async Task RemoveSelectedMemberAsync(CancellationToken cancellationToken)
    {
        if (SelectedRoleGroup == null || string.IsNullOrWhiteSpace(SelectedMember?.Identity))
        {
            return;
        }

        await UpdateMemberAsync(SelectedMember.Identity, RoleGroupMemberAction.Remove, cancellationToken);
    }

    private async Task UpdateMemberAsync(string member, RoleGroupMemberAction action, CancellationToken cancellationToken)
    {
        if (SelectedRoleGroup == null)
        {
            return;
        }

        var operation = action == RoleGroupMemberAction.Add
            ? "Add member to role group"
            : "Remove member from role group";
        if (!ConfirmMutation(
                operation,
                $"{member} -> {SelectedRoleGroup.Identity}",
                "Update the administrative membership of the role group.",
                "Confirm membership update"))
        {
            return;
        }

        IsSaving = true;
        ErrorMessage = null;

        try
        {
            var result = await _workerService.ModifyRoleGroupMemberAsync(
                new ModifyRoleGroupMemberRequest
                {
                    Identity = SelectedRoleGroup.Identity,
                    Member = member,
                    Action = action
                },
                cancellationToken: cancellationToken);

            if (!result.IsSuccess)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to update membership.";
                return;
            }

            _shellViewModel.AddLog(LogLevel.Information, $"{action} member {member} on {SelectedRoleGroup.Identity}", "Permissions");
            await RefreshAsync(cancellationToken);
            await LoadDetailsAsync(SelectedRoleGroup.Identity, cancellationToken);
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsSaving = false;
        }
    }

    private RoleGroupListItemDto? RestoreSelection(string? identity)
    {
        if (string.IsNullOrWhiteSpace(identity))
        {
            return RoleGroups.FirstOrDefault();
        }

        return RoleGroups.FirstOrDefault(group =>
            string.Equals(group.Identity, identity, StringComparison.OrdinalIgnoreCase))
            ?? RoleGroups.FirstOrDefault();
    }

    private void ResetEditor()
    {
        NewRoleGroupName = null;
        NewRoleGroupDescription = null;
        NewRoleGroupRoles = null;
        NewRoleGroupMembers = null;
        CopyFromRoleGroupIdentity = null;
    }

    private void ResetDisconnectedState()
    {
        RoleGroups.Clear();
        SelectedRoleGroup = null;
        SelectedRoleGroupDetails = null;
        TotalCount = 0;
        HasMore = false;
        ErrorMessage = "Not connected to Exchange Online";
    }

    private void RaiseCanExecuteChanged()
    {
        OnPropertyChanged(nameof(CanRefresh));
        OnPropertyChanged(nameof(CanLoadMore));
        OnPropertyChanged(nameof(CanSaveRoleGroup));
        OnPropertyChanged(nameof(CanAddMember));
        OnPropertyChanged(nameof(CanRemoveSelectedMember));
        CommandManager.InvalidateRequerySuggested();
    }

    private static List<string> ParseCsv(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return new List<string>();
        }

        return value
            .Split([',', ';', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static int GetRefreshPageSize(int loadedCount)
        => Math.Max(PageSize, loadedCount);

    private void HandleLoadProgressEvent(EventEnvelope evt)
    {
        _ = RunOnUiThreadAsync(() => LoadProgress.Apply(evt));
    }
}


