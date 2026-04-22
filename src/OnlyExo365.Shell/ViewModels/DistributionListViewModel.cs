using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Shell.ViewModels;

public class DistributionListViewModel : ViewModelBase
{
    private readonly NavigationService _navigationService;
    private readonly ShellViewModel _shellViewModel;
    private readonly DistributionListListViewModel _list;
    private readonly DistributionListDetailsViewModel _details;
    private readonly DistributionListSettingsEditorViewModel _settings;
    private readonly DistributionListEditorViewModel _editor;

    private string? _errorMessage;

    public DistributionListViewModel(IDistributionListsWorkerService workerService, NavigationService navigationService, ShellViewModel shellViewModel)
    {
        _navigationService = navigationService;
        _shellViewModel = shellViewModel;

        _list = new DistributionListListViewModel(workerService, shellViewModel, ViewDetails, SetErrorMessage);
        _details = new DistributionListDetailsViewModel(workerService, navigationService, shellViewModel, () => _list.SelectedItem, SetErrorMessage, OnDetailsChanged);
        _settings = new DistributionListSettingsEditorViewModel(workerService, shellViewModel, () => _details.SelectedDetails, SetErrorMessage);
        _editor = new DistributionListEditorViewModel(workerService, shellViewModel, SetErrorMessage, RefreshAsync);

        _list.PropertyChanged += OnChildPropertyChanged;
        _details.PropertyChanged += OnChildPropertyChanged;
        _settings.PropertyChanged += OnChildPropertyChanged;
        _editor.PropertyChanged += OnChildPropertyChanged;

        _navigationService.SelectedIdentityChanged += OnSelectedIdentityChanged;
        _shellViewModel.PropertyChanged += OnShellViewModelPropertyChanged;
    }

    public ObservableCollection<string> GroupTypeFilters => _list.GroupTypeFilters;
    public ObservableCollection<DistributionListItemDto> DistributionLists => _list.DistributionLists;
    public ObservableCollection<GroupMemberDto> Members => _details.Members;
    public ObservableCollection<string> AcceptMessagesOnlyFrom => _settings.AcceptMessagesOnlyFrom;
    public ObservableCollection<string> RejectMessagesFrom => _settings.RejectMessagesFrom;
    public ObservableCollection<string> AvailableMailDomains => _editor.AvailableMailDomains;

    public bool IsLoading => _list.IsLoading;
    public bool IsLoadingDetails => _details.IsLoadingDetails;
    public bool IsLoadingMembers => _details.IsLoadingMembers;
    public bool IsSavingSettings => _settings.IsSavingSettings;
    public bool IsBusy => IsLoading || IsLoadingDetails || IsLoadingMembers || IsSavingSettings || IsCreatingDistributionList;

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
        get => _list.SearchQuery;
        set => _list.SearchQuery = value;
    }

    public string SelectedGroupTypeFilter
    {
        get => _list.SelectedGroupTypeFilter;
        set => _list.SelectedGroupTypeFilter = value;
    }

    public bool IncludeDynamic
    {
        get => _list.IncludeDynamic;
        set => _list.IncludeDynamic = value;
    }

    public int TotalCount => _list.TotalCount;
    public bool HasMore => _list.HasMore;
    public bool MembersHasMore => _details.MembersHasMore;

    public DistributionListItemDto? SelectedItem
    {
        get => _list.SelectedItem;
        set => _list.SelectedItem = value;
    }

    public DistributionListDetailsDto? SelectedDetails => _details.SelectedDetails;
    public bool HasSelection => _details.HasSelection;
    public bool HasDetails => _details.HasDetails;
    public bool IsDynamicGroup => _details.IsDynamicGroup;
    public bool IsMicrosoft365Group => _details.IsMicrosoft365Group;
    public bool IsMailSecurityGroup => _details.IsMailSecurityGroup;
    public bool HasOwners => _details.HasOwners;
    public bool ShowMicrosoft365Metadata => _details.ShowMicrosoft365Metadata;
    public string CurrentGroupTypeLabel => _details.CurrentGroupTypeLabel;

    public string? NewMemberIdentity
    {
        get => _details.NewMemberIdentity;
        set => _details.NewMemberIdentity = value;
    }

    public string? NewAcceptedSender
    {
        get => _settings.NewAcceptedSender;
        set => _settings.NewAcceptedSender = value;
    }

    public string? NewRejectedSender
    {
        get => _settings.NewRejectedSender;
        set => _settings.NewRejectedSender = value;
    }

    public bool AllowExternalSenders
    {
        get => _settings.AllowExternalSenders;
        set => _settings.AllowExternalSenders = value;
    }

    public bool HasPendingSettingsChanges => _settings.HasPendingSettingsChanges;

    public string? NewDistributionListDisplayName
    {
        get => _editor.NewDistributionListDisplayName;
        set => _editor.NewDistributionListDisplayName = value;
    }

    public string? NewDistributionListAlias
    {
        get => _editor.NewDistributionListAlias;
        set => _editor.NewDistributionListAlias = value;
    }

    public string? NewDistributionListLocalPart
    {
        get => _editor.NewDistributionListLocalPart;
        set => _editor.NewDistributionListLocalPart = value;
    }

    public string? SelectedDistributionListDomain
    {
        get => _editor.SelectedDistributionListDomain;
        set => _editor.SelectedDistributionListDomain = value;
    }

    public bool IsCreatingDistributionList => _editor.IsCreatingDistributionList;

    public bool CanRefresh => _list.CanRefresh;
    public bool CanLoadMore => _list.CanLoadMore;
    public bool CanIncludeDynamicFilter => _list.CanIncludeDynamicFilter;
    public bool CanPreviewDynamicMembers => _details.CanPreviewDynamicMembers;
    public bool CanModifyMembers => _details.CanModifyMembers;
    public bool CanEditSettings => _settings.CanEditSettings;
    public bool CanEditExternalSenders => _settings.CanEditExternalSenders;
    public bool CanEditAcceptMessagesOnlyFrom => _settings.CanEditAcceptMessagesOnlyFrom;
    public bool CanEditRejectMessagesFrom => _settings.CanEditRejectMessagesFrom;
    public bool CanAddMember => _details.CanAddMember;
    public bool CanAddAcceptSender => _settings.CanAddAcceptSender;
    public bool CanAddRejectSender => _settings.CanAddRejectSender;
    public bool CanCreateDistributionList => _editor.CanCreateDistributionList;

    public string StatusText => _list.StatusText;
    public string MembersStatusText => _details.MembersStatusText;
    public string LoadingOverlayText => GetLoadingOverlayText();
    public CollectionLoadProgressViewModel LoadProgress => _list.LoadProgress;

    public ICommand RefreshCommand => _list.RefreshCommand;
    public ICommand LoadMoreCommand => _list.LoadMoreCommand;
    public ICommand CancelCommand => _list.CancelCommand;
    public ICommand ViewDetailsCommand => _list.ViewDetailsCommand;
    public ICommand BackCommand => _details.BackCommand;
    public ICommand LoadMoreMembersCommand => _details.LoadMoreMembersCommand;
    public ICommand PreviewDynamicMembersCommand => _details.PreviewDynamicMembersCommand;
    public ICommand AddMemberCommand => _details.AddMemberCommand;
    public ICommand RemoveMemberCommand => _details.RemoveMemberCommand;
    public ICommand SaveSettingsCommand => _settings.SaveSettingsCommand;
    public ICommand DiscardSettingsCommand => _settings.DiscardSettingsCommand;
    public ICommand AddAcceptSenderCommand => _settings.AddAcceptSenderCommand;
    public ICommand RemoveAcceptSenderCommand => _settings.RemoveAcceptSenderCommand;
    public ICommand AddRejectSenderCommand => _settings.AddRejectSenderCommand;
    public ICommand RemoveRejectSenderCommand => _settings.RemoveRejectSenderCommand;
    public ICommand CreateDistributionListCommand => _editor.CreateDistributionListCommand;

    public async Task LoadAsync()
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            await _list.LoadAsync(CancellationToken.None);
            return;
        }

        await _editor.LoadAcceptedDomainsAsync(CancellationToken.None);
        await _list.LoadAsync(CancellationToken.None);
    }

    public void Cancel() => _list.Cancel();

    private void OnSelectedIdentityChanged(object? sender, string? identity)
    {
        if (_navigationService.CurrentPage != NavigationPage.DistributionLists)
        {
            return;
        }

        if (!string.IsNullOrWhiteSpace(identity))
        {
            _ = _details.LoadDetailsAsync(identity, CancellationToken.None);
            return;
        }

        _details.ClearSelection();
    }

    private void OnShellViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        _list.HandleShellPropertyChanged(e);
        _details.HandleShellPropertyChanged();
        _settings.HandleShellPropertyChanged();
        _editor.HandleShellPropertyChanged();
    }

    private void OnChildPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (!string.IsNullOrWhiteSpace(e.PropertyName))
        {
            OnPropertyChanged(e.PropertyName);

            if (e.PropertyName is nameof(IsLoading) or nameof(IsLoadingDetails) or nameof(IsLoadingMembers) or nameof(IsSavingSettings) or nameof(IsCreatingDistributionList))
            {
                OnPropertyChanged(nameof(IsBusy));
                OnPropertyChanged(nameof(LoadingOverlayText));
            }
        }
    }

    private void ViewDetails(DistributionListItemDto? item)
    {
        if (item == null)
        {
            return;
        }

        _navigationService.NavigateToDetails(NavigationPage.DistributionLists, item.Identity, item);
    }

    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        await _list.LoadAsync(cancellationToken);
    }

    private void OnDetailsChanged(DistributionListDetailsDto? details)
    {
        _settings.InitializeFromDetails(details);
    }

    private void SetErrorMessage(string? errorMessage)
    {
        ErrorMessage = errorMessage;
    }

    private string GetLoadingOverlayText()
    {
        if (IsCreatingDistributionList)
        {
            return "Creating group...";
        }

        if (IsSavingSettings)
        {
            return "Saving settings group...";
        }

        if (IsLoadingMembers)
        {
            return "Loading group members...";
        }

        if (IsLoadingDetails)
        {
            return "Loading group details...";
        }

        return "Loading groups...";
    }
}

