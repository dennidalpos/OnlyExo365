using System.Collections.ObjectModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Paging;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

public class PublicFoldersViewModel : ViewModelBase
{
    private const NavigationPage AlertPage = NavigationPage.PublicFolders;
    private readonly IWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;
    private readonly DebounceHelper _searchDebounce = new();

    private bool _isLoading;
    private bool _isSaving;
    private bool _hasPendingChanges;
    private bool _suspendDirtyTracking;
    private string? _errorMessage;
    private string? _searchQuery;
    private string _mailEnabledFilter = "All";
    private int _totalCount;
    private int _currentSkip;
    private bool _hasMore;
    private PublicFolderListItemDto? _selectedFolder;

    private string? _identity;
    private string _name = string.Empty;
    private string _parentPath = "\\";
    private bool _mailEnabled;
    private string _alias = string.Empty;
    private string _primarySmtpAddress = string.Empty;
    private bool _hiddenFromAddressListsEnabled;
    private bool _hasSubFolders;
    private int? _itemCount;
    private string _totalItemSize = string.Empty;
    private string? _newPermissionUser;
    private string _newPermissionRole = "Editor";
    private string _permissionsSummary = string.Empty;

    private const int PageSize = PagingDefaults.DefaultPageSize;

    public PublicFoldersViewModel(IWorkerService workerService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;

        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
        LoadMoreCommand = new AsyncRelayCommand(LoadMoreAsync, () => CanLoadMore);
        SaveCommand = new AsyncRelayCommand(SaveAsync, () => CanSave);
        NewFolderCommand = new RelayCommand(BeginCreate, () => !IsSaving);
        AddPermissionCommand = new AsyncRelayCommand(AddPermissionAsync, () => CanAddPermission);
        UpdatePermissionCommand = new AsyncRelayCommand<object>(UpdatePermissionAsync, CanUpdatePermission);
        RemovePermissionCommand = new AsyncRelayCommand<object>(RemovePermissionAsync, CanRemovePermission);
        DeleteCommand = new AsyncRelayCommand(DeleteAsync, () => CanDelete);
    }

    public ObservableCollection<PublicFolderListItemDto> Folders { get; } = new();
    public ObservableCollection<PublicFolderPermissionDisplayItem> PermissionEntries { get; } = new();
    public CollectionLoadProgressViewModel LoadProgress { get; } = new();

    public IReadOnlyList<string> MailEnabledFilters { get; } = new[] { "All", "Mail-enabled only", "Mail-disabled only" };
    public IReadOnlyList<string> PublicFolderPermissionRoles { get; } =
    [
        "Owner",
        "PublishingEditor",
        "Editor",
        "PublishingAuthor",
        "Author",
        "NonEditingAuthor",
        "Reviewer",
        "Contributor"
    ];

    public bool IsLoading
    {
        get => _isLoading;
        private set
        {
            if (SetProperty(ref _isLoading, value))
            {
                OnPropertyChanged(nameof(CanRefresh));
                OnPropertyChanged(nameof(CanLoadMore));
                OnPropertyChanged(nameof(CanSave));
                OnPropertyChanged(nameof(CanDelete));
                OnPropertyChanged(nameof(CanAddPermission));
                CommandManager.InvalidateRequerySuggested();
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
                OnPropertyChanged(nameof(CanSave));
                OnPropertyChanged(nameof(CanDelete));
                OnPropertyChanged(nameof(CanAddPermission));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool HasPendingChanges
    {
        get => _hasPendingChanges;
        set => SetProperty(ref _hasPendingChanges, value);
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

    public string MailEnabledFilter
    {
        get => _mailEnabledFilter;
        set
        {
            if (SetProperty(ref _mailEnabledFilter, value))
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
                OnPropertyChanged(nameof(CanLoadMore));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string StatusText => $"{Folders.Count} of {TotalCount} public folders";

    public PublicFolderListItemDto? SelectedFolder
    {
        get => _selectedFolder;
        set
        {
            if (SetProperty(ref _selectedFolder, value) && value != null)
            {
                _ = LoadDetailsAsync(value);
            }
        }
    }

    public string? Identity
    {
        get => _identity;
        private set
        {
            if (SetProperty(ref _identity, value))
            {
                OnPropertyChanged(nameof(IsEditingExistingFolder));
                OnPropertyChanged(nameof(CanDelete));
                OnPropertyChanged(nameof(CanAddPermission));
            }
        }
    }

    public bool IsEditingExistingFolder => !string.IsNullOrWhiteSpace(Identity);

    public string Name
    {
        get => _name;
        set => SetTrackedProperty(ref _name, value, nameof(CanSave));
    }

    public string ParentPath
    {
        get => _parentPath;
        set => SetTrackedProperty(ref _parentPath, NormalizeParentPath(value));
    }

    public bool MailEnabled
    {
        get => _mailEnabled;
        set => SetTrackedProperty(ref _mailEnabled, value, nameof(CanSave));
    }

    public string Alias
    {
        get => _alias;
        set => SetTrackedProperty(ref _alias, value, nameof(CanSave));
    }

    public string PrimarySmtpAddress
    {
        get => _primarySmtpAddress;
        set => SetTrackedProperty(ref _primarySmtpAddress, value);
    }

    public bool HiddenFromAddressListsEnabled
    {
        get => _hiddenFromAddressListsEnabled;
        set => SetTrackedProperty(ref _hiddenFromAddressListsEnabled, value);
    }

    public bool HasSubFolders
    {
        get => _hasSubFolders;
        private set => SetProperty(ref _hasSubFolders, value);
    }

    public int? ItemCount
    {
        get => _itemCount;
        private set => SetProperty(ref _itemCount, value);
    }

    public string TotalItemSize
    {
        get => _totalItemSize;
        private set => SetProperty(ref _totalItemSize, value);
    }

    public string PermissionsSummary
    {
        get => _permissionsSummary;
        private set => SetProperty(ref _permissionsSummary, value);
    }

    public string? NewPermissionUser
    {
        get => _newPermissionUser;
        set
        {
            if (SetProperty(ref _newPermissionUser, value))
            {
                OnPropertyChanged(nameof(CanAddPermission));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string NewPermissionRole
    {
        get => _newPermissionRole;
        set
        {
            if (SetProperty(ref _newPermissionRole, value))
            {
                OnPropertyChanged(nameof(CanAddPermission));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool CanRefresh => !IsLoading && _shellViewModel.IsExchangeConnected;
    public bool CanLoadMore => !IsLoading && HasMore && _shellViewModel.IsExchangeConnected;
    public bool CanDelete => IsEditingExistingFolder && !IsLoading && !IsSaving && _shellViewModel.IsExchangeConnected;
    public bool CanAddPermission =>
        IsEditingExistingFolder &&
        !IsLoading &&
        !IsSaving &&
        _shellViewModel.IsExchangeConnected &&
        !string.IsNullOrWhiteSpace(NewPermissionUser) &&
        !string.IsNullOrWhiteSpace(NewPermissionRole);

    public bool CanSave =>
        !IsLoading &&
        !IsSaving &&
        _shellViewModel.IsExchangeConnected &&
        !string.IsNullOrWhiteSpace(Name) &&
        (!MailEnabled || !string.IsNullOrWhiteSpace(Alias));

    public ICommand RefreshCommand { get; }
    public ICommand LoadMoreCommand { get; }
    public ICommand SaveCommand { get; }
    public ICommand NewFolderCommand { get; }
    public ICommand AddPermissionCommand { get; }
    public ICommand UpdatePermissionCommand { get; }
    public ICommand RemovePermissionCommand { get; }
    public ICommand DeleteCommand { get; }

    public async Task LoadAsync()
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            Folders.Clear();
            ErrorMessage = null;
            _shellViewModel.ClearPageAlert(AlertPage);
            return;
        }

        if (Folders.Count == 0)
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
            _shellViewModel.AddLog(LogLevel.Error, $"Public folders refresh failed: {ex.Message}", "PublicFolders");
        }
    }

    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        var hasExistingFolders = Folders.Count > 0;
        if (!_shellViewModel.IsExchangeConnected)
        {
            Folders.Clear();
            ErrorMessage = null;
            TotalCount = 0;
            HasMore = false;
            _shellViewModel.ClearPageAlert(AlertPage);
            return;
        }

        IsLoading = true;
        LoadProgress.Start("Loading public folders...", "folders");
        ErrorMessage = null;
        _shellViewModel.ClearPageAlert(AlertPage);
        var refreshPageSize = GetRefreshPageSize(Folders.Count);
        _currentSkip = 0;

        try
        {
            var result = await _workerService.GetPublicFoldersAsync(
                new GetPublicFoldersRequest
                {
                    SearchQuery = string.IsNullOrWhiteSpace(SearchQuery) ? null : SearchQuery.Trim(),
                    MailEnabledOnly = NormalizeMailEnabledFilter(MailEnabledFilter),
                    PageSize = refreshPageSize,
                    Skip = 0
                },
                eventHandler: HandleLoadProgressEvent,
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                var errorMessage = result.Error?.Message ?? "Unable to load public folders";
                ErrorMessage = hasExistingFolders ? errorMessage : null;
                if (!hasExistingFolders)
                {
                    _shellViewModel.ShowPageLoadFailedAlert(AlertPage, errorMessage);
                }
                return;
            }

            Folders.ReplaceAll(result.Value.Folders);
            TotalCount = result.Value.TotalCount;
            HasMore = result.Value.HasMore;
            _currentSkip = Folders.Count;
            _shellViewModel.ClearPageAlert(AlertPage);
        }
        catch (Exception ex)
        {
            ErrorMessage = hasExistingFolders ? ex.Message : null;
            if (!hasExistingFolders)
            {
                _shellViewModel.ShowPageLoadFailedAlert(AlertPage, ex.Message);
            }
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
        LoadProgress.Start("Loading public folders...", "folders");
        ErrorMessage = null;

        try
        {
            var result = await _workerService.GetPublicFoldersAsync(
                new GetPublicFoldersRequest
                {
                    SearchQuery = string.IsNullOrWhiteSpace(SearchQuery) ? null : SearchQuery.Trim(),
                    MailEnabledOnly = NormalizeMailEnabledFilter(MailEnabledFilter),
                    PageSize = PageSize,
                    Skip = _currentSkip
                },
                eventHandler: HandleLoadProgressEvent,
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to load public folders";
                return;
            }

            foreach (var item in result.Value.Folders)
            {
                Folders.Add(item);
            }

            TotalCount = result.Value.TotalCount;
            HasMore = result.Value.HasMore;
            _currentSkip = Folders.Count;
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

    private async Task LoadDetailsAsync(PublicFolderListItemDto selected)
    {
        ErrorMessage = null;

        try
        {
            var result = await _workerService.GetPublicFolderDetailsAsync(
                new GetPublicFolderDetailsRequest
                {
                    Identity = selected.Identity
                },
                cancellationToken: CancellationToken.None);

            if (!result.IsSuccess || result.Value == null)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to load public folder details.";
                return;
            }

            ApplyDetails(result.Value);
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
    }

    private async Task SaveAsync(CancellationToken cancellationToken)
    {
        if (!CanSave)
        {
            return;
        }

        if (!ConfirmMutation(
                string.IsNullOrWhiteSpace(Identity) ? "Creating public folder" : "Updating public folder",
                string.IsNullOrWhiteSpace(Identity) ? Name.Trim() : Identity,
                "Save the public folder structure, mail settings, and attributes.",
                "Confirm public folder save"))
        {
            return;
        }

        IsSaving = true;
        ErrorMessage = null;

        try
        {
            var result = await _workerService.UpsertPublicFolderAsync(
                new UpsertPublicFolderRequest
                {
                    Identity = Identity,
                    Name = Name.Trim(),
                    ParentPath = NormalizeParentPath(ParentPath),
                    MailEnabled = MailEnabled,
                    Alias = string.IsNullOrWhiteSpace(Alias) ? null : Alias.Trim(),
                    PrimarySmtpAddress = string.IsNullOrWhiteSpace(PrimarySmtpAddress) ? null : PrimarySmtpAddress.Trim(),
                    HiddenFromAddressListsEnabled = HiddenFromAddressListsEnabled
                },
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to save the public folder.";
                return;
            }

            _shellViewModel.AddLog(LogLevel.Information, $"Public folder saved: {result.Value.Identity}", "PublicFolders");
            HasPendingChanges = false;
            await RefreshAsync(cancellationToken);
            await TryReselectSavedFolderAsync(result.Value.Identity);
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

    private async Task TryReselectSavedFolderAsync(string identity)
    {
        var match = Folders.FirstOrDefault(folder =>
            string.Equals(folder.Identity, identity, StringComparison.OrdinalIgnoreCase));

        if (match != null)
        {
            SelectedFolder = match;
            await LoadDetailsAsync(match);
            HasPendingChanges = false;
        }
    }

    private void BeginCreate()
    {
        _suspendDirtyTracking = true;

        try
        {
            SelectedFolder = null;
            Identity = null;
            Name = string.Empty;
            ParentPath = "\\";
            MailEnabled = false;
            Alias = string.Empty;
            PrimarySmtpAddress = string.Empty;
            HiddenFromAddressListsEnabled = false;
            HasSubFolders = false;
            ItemCount = null;
            TotalItemSize = string.Empty;
            PermissionEntries.Clear();
            PermissionsSummary = string.Empty;
            NewPermissionUser = string.Empty;
            NewPermissionRole = "Editor";
            ErrorMessage = null;
        }
        finally
        {
            _suspendDirtyTracking = false;
            HasPendingChanges = false;
        }
    }

    private void ApplyDetails(PublicFolderDetailsDto details)
    {
        _suspendDirtyTracking = true;

        try
        {
            Identity = details.Identity;
            Name = details.Name;
            ParentPath = details.ParentPath;
            MailEnabled = details.MailEnabled;
            Alias = details.Alias ?? string.Empty;
            PrimarySmtpAddress = details.PrimarySmtpAddress ?? string.Empty;
            HiddenFromAddressListsEnabled = details.HiddenFromAddressListsEnabled;
            HasSubFolders = details.HasSubFolders;
            ItemCount = details.ItemCount;
            TotalItemSize = details.TotalItemSize ?? string.Empty;
            PermissionEntries.Clear();
            foreach (var permission in details.Permissions
                         .OrderBy(static item => item.DisplayName, StringComparer.OrdinalIgnoreCase)
                         .ThenBy(static item => item.User, StringComparer.OrdinalIgnoreCase))
            {
                PermissionEntries.Add(new PublicFolderPermissionDisplayItem
                {
                    User = permission.User,
                    DisplayName = string.IsNullOrWhiteSpace(permission.DisplayName) ? permission.User : permission.DisplayName,
                    SelectedRole = permission.AccessRights.FirstOrDefault() ?? string.Join(", ", permission.AccessRights),
                    AccessRightsSummary = permission.AccessRights.Count == 0
                        ? string.Empty
                        : string.Join(", ", permission.AccessRights)
                });
            }
            PermissionsSummary = details.Permissions.Count == 0
                ? "No explicit client permissions detected"
                : string.Join(Environment.NewLine, details.Permissions.Select(permission =>
                    $"{permission.User}: {string.Join(", ", permission.AccessRights)}"));
        }
        finally
        {
            _suspendDirtyTracking = false;
            HasPendingChanges = false;
        }
    }

    private void SetTrackedProperty<T>(ref T field, T value, params string[] additionalProperties)
    {
        if (!SetProperty(ref field, value))
        {
            return;
        }

        foreach (var property in additionalProperties)
        {
            OnPropertyChanged(property);
        }

        if (!_suspendDirtyTracking)
        {
            HasPendingChanges = true;
        }
    }

    private static bool? NormalizeMailEnabledFilter(string? value)
    {
        return value switch
        {
            "Mail-enabled only" => true,
            "Mail-disabled only" => false,
            _ => null
        };
    }

    private static string NormalizeParentPath(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "\\";
        }

        var trimmed = value.Trim();
        return trimmed.StartsWith('\\') ? trimmed : "\\" + trimmed;
    }

    private static int GetRefreshPageSize(int loadedCount)
        => Math.Max(PageSize, loadedCount);

    private async Task AddPermissionAsync(CancellationToken cancellationToken)
    {
        if (!CanAddPermission || string.IsNullOrWhiteSpace(Identity))
        {
            return;
        }

        var normalizedUser = NewPermissionUser!.Trim();
        var existing = PermissionEntries.FirstOrDefault(entry =>
            string.Equals(entry.User, normalizedUser, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(entry.DisplayName, normalizedUser, StringComparison.OrdinalIgnoreCase));

        await ApplyPermissionChangeAsync(
            normalizedUser,
            existing == null ? PermissionAction.Add : PermissionAction.Modify,
            NewPermissionRole,
            cancellationToken);

        NewPermissionUser = string.Empty;
    }

    private async Task UpdatePermissionAsync(object? parameter, CancellationToken cancellationToken)
    {
        if (parameter is not PublicFolderPermissionDisplayItem entry)
        {
            return;
        }

        await ApplyPermissionChangeAsync(entry.User, PermissionAction.Modify, entry.SelectedRole, cancellationToken);
    }

    private bool CanUpdatePermission(object? parameter)
        => parameter is PublicFolderPermissionDisplayItem entry &&
           !string.IsNullOrWhiteSpace(entry.SelectedRole) &&
           IsEditingExistingFolder &&
           !IsLoading &&
           !IsSaving &&
           _shellViewModel.IsExchangeConnected;

    private async Task RemovePermissionAsync(object? parameter, CancellationToken cancellationToken)
    {
        if (parameter is not PublicFolderPermissionDisplayItem entry)
        {
            return;
        }

        await ApplyPermissionChangeAsync(entry.User, PermissionAction.Remove, null, cancellationToken);
    }

    private bool CanRemovePermission(object? parameter)
        => parameter is PublicFolderPermissionDisplayItem &&
           IsEditingExistingFolder &&
           !IsLoading &&
           !IsSaving &&
           _shellViewModel.IsExchangeConnected;

    private async Task ApplyPermissionChangeAsync(
        string user,
        PermissionAction action,
        string? role,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(Identity))
        {
            return;
        }

        var operation = action switch
        {
            PermissionAction.Add => "Add client permission",
            PermissionAction.Modify => "Updating client permission",
            PermissionAction.Remove => "Remove client permission",
            _ => "Updating client permission"
        };
        var impact = action == PermissionAction.Remove
            ? "The delegate loses explicit client access to the public folder."
            : $"The delegate receives the '{role ?? "N/A"}' role on the public folder.";
        if (!ConfirmMutation(
                operation,
                $"{user} -> {Identity}",
                impact,
                "Confirm public folder permission update"))
        {
            return;
        }

        IsSaving = true;
        ErrorMessage = null;

        try
        {
            var result = await _workerService.SetPublicFolderClientPermissionAsync(
                new SetPublicFolderClientPermissionRequest
                {
                    Identity = Identity,
                    User = user,
                    Action = action,
                    AccessRights = string.IsNullOrWhiteSpace(role) ? [] : [role]
                },
                cancellationToken: cancellationToken);

            if (!result.IsSuccess)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to update public folder client permissions.";
                return;
            }

            _shellViewModel.AddLog(LogLevel.Information, $"{action} public folder permission for {user} on {Identity}", "PublicFolders");

            var selected = SelectedFolder;
            if (selected != null)
            {
                await LoadDetailsAsync(selected);
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
            IsSaving = false;
        }
    }

    private async Task DeleteAsync(CancellationToken cancellationToken)
    {
        if (!CanDelete || string.IsNullOrWhiteSpace(Identity))
        {
            return;
        }

        if (!ConfirmMutation(
                HasSubFolders ? "Deleting recursive public folder" : "Deleting public folder",
                Identity,
                HasSubFolders
                    ? "Remove the public folder and all of its subfolders."
                    : "Permanently remove the selected public folder.",
                "Confirm public folder deletion"))
        {
            return;
        }

        IsSaving = true;
        ErrorMessage = null;

        try
        {
            var result = await _workerService.RemovePublicFolderAsync(
                new RemovePublicFolderRequest
                {
                    Identity = Identity,
                    Recursive = HasSubFolders
                },
                cancellationToken: cancellationToken);

            if (!result.IsSuccess)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to delete the public folder";
                return;
            }

            _shellViewModel.AddLog(LogLevel.Warning, $"Public folder deleted: {Identity}", "PublicFolders");
            BeginCreate();
            await RefreshAsync(cancellationToken);
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
            IsSaving = false;
        }
    }

    private void HandleLoadProgressEvent(EventEnvelope evt)
    {
        _ = RunOnUiThreadAsync(() => LoadProgress.Apply(evt));
    }
}

public sealed class PublicFolderPermissionDisplayItem : ViewModelBase
{
    private string _selectedRole = string.Empty;

    public string User { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public string AccessRightsSummary { get; set; } = string.Empty;

    public string SelectedRole
    {
        get => _selectedRole;
        set => SetProperty(ref _selectedRole, value);
    }
}
