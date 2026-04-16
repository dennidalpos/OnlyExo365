using System.Linq;
using System.Collections.ObjectModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Paging;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

public sealed class DistributionListDetailsViewModel : ViewModelBase
{
    private const int PageSize = PagingDefaults.DefaultPageSize;

    private readonly IDistributionListsWorkerService _workerService;
    private readonly NavigationService _navigationService;
    private readonly ShellViewModel _shellViewModel;
    private readonly Func<DistributionListItemDto?> _getSelectedItem;
    private readonly Action<string?> _setErrorMessage;
    private readonly Action<DistributionListDetailsDto?> _onDetailsChanged;

    private bool _isLoadingDetails;
    private bool _isLoadingMembers;
    private bool _membersHasMore;
    private int _membersCurrentSkip;
    private DistributionListDetailsDto? _selectedDetails;
    private string? _newMemberIdentity;

    public DistributionListDetailsViewModel(
        IDistributionListsWorkerService workerService,
        NavigationService navigationService,
        ShellViewModel shellViewModel,
        Func<DistributionListItemDto?> getSelectedItem,
        Action<string?> setErrorMessage,
        Action<DistributionListDetailsDto?> onDetailsChanged)
    {
        _workerService = workerService;
        _navigationService = navigationService;
        _shellViewModel = shellViewModel;
        _getSelectedItem = getSelectedItem;
        _setErrorMessage = setErrorMessage;
        _onDetailsChanged = onDetailsChanged;

        BackCommand = new RelayCommand(GoBack);
        LoadMoreMembersCommand = new AsyncRelayCommand(LoadMoreMembersAsync, () => MembersHasMore && !IsLoadingMembers);
        PreviewDynamicMembersCommand = new AsyncRelayCommand(PreviewDynamicMembersAsync, () => CanPreviewDynamicMembers);
        AddMemberCommand = new AsyncRelayCommand(AddMemberAsync, () => CanAddMember);
        RemoveMemberCommand = new AsyncRelayCommand<GroupMemberDto>(RemoveMemberAsync, member => member != null && CanModifyMembers);
    }

    public ObservableCollection<GroupMemberDto> Members { get; } = new();

    public bool IsLoadingDetails
    {
        get => _isLoadingDetails;
        private set => SetProperty(ref _isLoadingDetails, value);
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

    public DistributionListDetailsDto? SelectedDetails
    {
        get => _selectedDetails;
        private set
        {
            if (SetProperty(ref _selectedDetails, value))
            {
                OnPropertyChanged(nameof(HasSelection));
                OnPropertyChanged(nameof(HasDetails));
                OnPropertyChanged(nameof(IsDynamicGroup));
                OnPropertyChanged(nameof(IsMicrosoft365Group));
                OnPropertyChanged(nameof(IsMailSecurityGroup));
                OnPropertyChanged(nameof(CanModifyMembers));
                OnPropertyChanged(nameof(CanPreviewDynamicMembers));
                OnPropertyChanged(nameof(HasOwners));
                OnPropertyChanged(nameof(ShowMicrosoft365Metadata));
                OnPropertyChanged(nameof(CurrentGroupTypeLabel));
                _onDetailsChanged(value);
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool HasSelection => _getSelectedItem() != null || SelectedDetails != null;
    public bool HasDetails => SelectedDetails != null;
    public bool IsDynamicGroup => string.Equals(SelectedDetails?.GroupType, "Dynamic", StringComparison.OrdinalIgnoreCase);
    public bool IsMicrosoft365Group => string.Equals(SelectedDetails?.GroupType, "Microsoft365", StringComparison.OrdinalIgnoreCase);
    public bool IsMailSecurityGroup => string.Equals(SelectedDetails?.GroupType, "MailSecurity", StringComparison.OrdinalIgnoreCase);
    public bool HasOwners => SelectedDetails?.Owners?.Members.Count > 0;
    public bool ShowMicrosoft365Metadata => IsMicrosoft365Group;
    public string CurrentGroupTypeLabel => DistributionListViewModelSupport.FormatGroupTypeLabel(SelectedDetails?.GroupType);

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

    public bool CanPreviewDynamicMembers => IsDynamicGroup &&
                                            !IsLoadingMembers &&
                                            _shellViewModel.IsFeatureAvailable(
                                                f => f.CanGetDynamicDistributionGroup &&
                                                     (f.CanGetDynamicDistributionGroupMember || f.CanGetRecipient));
    public bool CanModifyMembers => HasDetails && !IsDynamicGroup && (IsMicrosoft365Group
        ? _shellViewModel.IsFeatureAvailable(f => f.CanAddUnifiedGroupLinks) || _shellViewModel.IsFeatureAvailable(f => f.CanRemoveUnifiedGroupLinks)
        : _shellViewModel.IsFeatureAvailable(f => f.CanAddDistributionGroupMember) || _shellViewModel.IsFeatureAvailable(f => f.CanRemoveDistributionGroupMember));
    public bool CanAddMember => !string.IsNullOrWhiteSpace(NewMemberIdentity) && CanModifyMembers;
    public string MembersStatusText => SelectedDetails?.Members != null
        ? $"{Members.Count} of {SelectedDetails.Members.TotalCount} members"
        : $"{Members.Count} members";

    public ICommand BackCommand { get; }
    public ICommand LoadMoreMembersCommand { get; }
    public ICommand PreviewDynamicMembersCommand { get; }
    public ICommand AddMemberCommand { get; }
    public ICommand RemoveMemberCommand { get; }

    public void HandleShellPropertyChanged()
    {
        OnPropertyChanged(nameof(CanModifyMembers));
        OnPropertyChanged(nameof(CanPreviewDynamicMembers));
        CommandManager.InvalidateRequerySuggested();
    }

    public async Task LoadDetailsAsync(string identity, CancellationToken cancellationToken)
    {
        var preserveLoadedMembers = SelectedDetails != null &&
            string.Equals(SelectedDetails.Identity, identity, StringComparison.OrdinalIgnoreCase);
        var refreshPageSize = GetRefreshPageSize(preserveLoadedMembers ? Members.Count : 0);

        IsLoadingDetails = true;
        _setErrorMessage(null);
        if (!preserveLoadedMembers)
        {
            Members.Clear();
            _membersCurrentSkip = 0;
        }

        try
        {
            var request = new GetDistributionListDetailsRequest
            {
                Identity = identity,
                GroupTypeHint = _getSelectedItem()?.GroupType ?? SelectedDetails?.GroupType,
                IncludeMembers = true,
                MembersPageSize = refreshPageSize
            };

            var result = await _workerService.GetDistributionListDetailsAsync(request, cancellationToken: cancellationToken);
            if (result.IsSuccess && result.Value != null)
            {
                SelectedDetails = result.Value;
                if (result.Value.Members != null)
                {
                    Members.ReplaceAll(result.Value.Members.Members);
                    MembersHasMore = result.Value.Members.HasMore;
                    _membersCurrentSkip = Members.Count;
                }
                else
                {
                    MembersHasMore = false;
                }

                OnPropertyChanged(nameof(MembersStatusText));
                return;
            }

            if (!result.WasCancelled)
            {
                var errorMessage = result.Error?.Message ?? "Unable to load distribution list details.";
                _setErrorMessage(errorMessage);
                _shellViewModel.AddLog(LogLevel.Error, $"Group details load failed: {errorMessage}");
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            _setErrorMessage(ex.Message);
            _shellViewModel.AddLog(LogLevel.Error, $"Group details error: {ex.Message}");
        }
        finally
        {
            IsLoadingDetails = false;
        }
    }

    public void ClearSelection()
    {
        SelectedDetails = null;
        Members.Clear();
        MembersHasMore = false;
        _membersCurrentSkip = 0;
        NewMemberIdentity = string.Empty;
        OnPropertyChanged(nameof(MembersStatusText));
    }

    private void GoBack()
    {
        ClearSelection();
        _navigationService.ClearSelection();
    }

    private async Task LoadMoreMembersAsync(CancellationToken cancellationToken)
    {
        if (!MembersHasMore || IsLoadingMembers || SelectedDetails == null)
        {
            return;
        }

        IsLoadingMembers = true;
        try
        {
            var request = new GetGroupMembersRequest
            {
                Identity = SelectedDetails.Identity,
                GroupType = DistributionListViewModelSupport.MapGroupTypeForWorker(SelectedDetails.GroupType),
                PageSize = PageSize,
                Skip = _membersCurrentSkip
            };

            var result = await _workerService.GetGroupMembersAsync(request, cancellationToken: cancellationToken);
            if (result.IsSuccess && result.Value != null)
            {
                foreach (var member in result.Value.Members)
                {
                    Members.Add(member);
                }

                UpdateMembersPageState(result.Value.TotalCount, result.Value.HasMore, result.Value.PageSize, Members.Count);
                MembersHasMore = result.Value.HasMore;
                _membersCurrentSkip = Members.Count;
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
        if (!IsDynamicGroup || SelectedDetails == null)
        {
            return;
        }

        IsLoadingMembers = true;
        Members.Clear();

        try
        {
            var result = await _workerService.PreviewDynamicGroupMembersAsync(
                new PreviewDynamicGroupMembersRequest
                {
                    Identity = SelectedDetails.Identity,
                    MaxResults = PageSize
                },
                cancellationToken: cancellationToken);

            if (result.IsSuccess && result.Value != null)
            {
                foreach (var member in result.Value.Members)
                {
                    Members.Add(member);
                }

                var hasMore = result.Value.TotalCount > result.Value.Members.Count;
                SelectedDetails.Members ??= new GroupMembersPageDto();
                SelectedDetails.Members.Members = result.Value.Members.ToList();
                SelectedDetails.Members.TotalCount = result.Value.TotalCount;
                SelectedDetails.Members.PageSize = PageSize;
                SelectedDetails.Members.Skip = 0;
                SelectedDetails.Members.HasMore = hasMore;
                MembersHasMore = hasMore;
                _membersCurrentSkip = Members.Count;
                OnPropertyChanged(nameof(MembersStatusText));

                if (!string.IsNullOrWhiteSpace(result.Value.Warning))
                {
                    _shellViewModel.AddLog(LogLevel.Warning, result.Value.Warning);
                }

                _shellViewModel.AddLog(LogLevel.Information, $"Anteprima completata: {result.Value.Members.Count} of {result.Value.TotalCount} members");
            }
            else if (!result.WasCancelled)
            {
                _shellViewModel.AddLog(LogLevel.Error, $"Preview members failed: {result.Error?.Message}");
            }
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Preview members error: {ex.Message}");
        }
        finally
        {
            IsLoadingMembers = false;
        }
    }

    private async Task AddMemberAsync(CancellationToken cancellationToken)
    {
        if (!CanAddMember || SelectedDetails == null)
        {
            return;
        }

        try
        {
            var memberIdentity = NewMemberIdentity!.Trim();
            if (!ConfirmMutation(
                    "Add member to distribution list",
                    $"{memberIdentity} -> {SelectedDetails.Identity}",
                    "Refresh the membership of the selected distribution list.",
                    "Confirm membership update"))
            {
                return;
            }

            var result = await _workerService.ModifyGroupMemberAsync(
                new ModifyGroupMemberRequest
                {
                    Identity = SelectedDetails.Identity,
                    Member = memberIdentity,
                    Action = GroupMemberAction.Add,
                    GroupType = DistributionListViewModelSupport.MapGroupTypeForWorker(SelectedDetails.GroupType)
                },
                cancellationToken: cancellationToken);

            if (result.IsSuccess)
            {
                NewMemberIdentity = string.Empty;
                await LoadDetailsAsync(SelectedDetails.Identity, cancellationToken);
            }
            else if (!result.WasCancelled)
            {
                _shellViewModel.AddLog(LogLevel.Error, $"Add member failed: {result.Error?.Message}");
            }
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Add member error: {ex.Message}");
        }
    }

    private async Task RemoveMemberAsync(GroupMemberDto? member, CancellationToken cancellationToken)
    {
        if (member == null || SelectedDetails == null || !CanModifyMembers)
        {
            return;
        }

        try
        {
            if (!ConfirmMutation(
                    "Remove member from distribution list",
                    $"{member.Identity} <- {SelectedDetails.Identity}",
                    "The member is removed from the selected distribution list.",
                    "Confirm membership update"))
            {
                return;
            }

            var result = await _workerService.ModifyGroupMemberAsync(
                new ModifyGroupMemberRequest
                {
                    Identity = SelectedDetails.Identity,
                    Member = member.Identity,
                    Action = GroupMemberAction.Remove,
                    GroupType = DistributionListViewModelSupport.MapGroupTypeForWorker(SelectedDetails.GroupType)
                },
                cancellationToken: cancellationToken);

            if (result.IsSuccess)
            {
                Members.Remove(member);
                OnPropertyChanged(nameof(MembersStatusText));
            }
            else if (!result.WasCancelled)
            {
                _shellViewModel.AddLog(LogLevel.Error, $"Remove member failed: {result.Error?.Message}");
            }
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Remove member error: {ex.Message}");
        }
    }

    private void UpdateMembersPageState(int totalCount, bool hasMore, int pageSize, int skip)
    {
        if (SelectedDetails?.Members == null)
        {
            return;
        }

        SelectedDetails.Members.TotalCount = totalCount;
        SelectedDetails.Members.HasMore = hasMore;
        SelectedDetails.Members.PageSize = pageSize;
        SelectedDetails.Members.Skip = skip;
    }

    private static int GetRefreshPageSize(int loadedCount)
        => Math.Max(PageSize, loadedCount);
}
