using System.Collections.ObjectModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;

namespace ExchangeAdmin.Presentation.ViewModels;

public sealed class MailboxPermissionsEditorViewModel : ViewModelBase
{
    private readonly IMailboxesWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;
    private readonly Func<string?> _getMailboxIdentity;
    private readonly Func<string?> _getPrimarySmtpAddress;
    private readonly Func<bool> _getIsSaving;
    private readonly Action<bool> _setIsSaving;
    private readonly Action<string?> _setErrorMessage;
    private readonly Func<CancellationToken, Task> _refreshAsync;

    private readonly ObservableCollection<MailboxPermissionChangeItem> _pendingActions = new();

    private bool _hasPendingChanges;
    private string? _newPermissionUser;
    private PermissionType _newPermissionType = PermissionType.FullAccess;
    private bool _newPermissionAutoMapping = true;
    private string _folderPermissionFolderPath = "Calendar";
    private string? _newFolderPermissionUser;
    private string _newFolderPermissionRole = "Editor";
    private bool _isFolderPermissionsLoading;
    private bool _isFolderPermissionSaving;
    private string _folderPermissionTargetLabel = "Calendar";
    private MailboxPermissionsDto? _permissions;

    public MailboxPermissionsEditorViewModel(
        IMailboxesWorkerService workerService,
        ShellViewModel shellViewModel,
        Func<string?> getMailboxIdentity,
        Func<string?> getPrimarySmtpAddress,
        Func<bool> getIsSaving,
        Action<bool> setIsSaving,
        Action<string?> setErrorMessage,
        Func<CancellationToken, Task> refreshAsync)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;
        _getMailboxIdentity = getMailboxIdentity;
        _getPrimarySmtpAddress = getPrimarySmtpAddress;
        _getIsSaving = getIsSaving;
        _setIsSaving = setIsSaving;
        _setErrorMessage = setErrorMessage;
        _refreshAsync = refreshAsync;

        SavePermissionsCommand = new AsyncRelayCommand(SavePermissionsAsync, () => HasPendingChanges && !_getIsSaving());
        DiscardPermissionsCommand = new RelayCommand(DiscardPendingChanges, () => HasPendingChanges);
        AddPermissionCommand = new RelayCommand(AddPermission, () => CanAddPermission);
        RemovePermissionCommand = new RelayCommand<object>(RemovePermission);
        ModifyAutoMappingCommand = new RelayCommand<object>(ModifyAutoMapping);
        RefreshFolderPermissionsCommand = new AsyncRelayCommand(LoadFolderPermissionsAsync, () => CanLoadFolderPermissions);
        AddFolderPermissionCommand = new AsyncRelayCommand(AddOrUpdateFolderPermissionAsync, () => CanAddFolderPermission);
        UpdateFolderPermissionCommand = new AsyncRelayCommand<object>(UpdateFolderPermissionAsync, CanUpdateFolderPermission);
        RemoveFolderPermissionCommand = new AsyncRelayCommand<object>(RemoveFolderPermissionAsync, CanRemoveFolderPermission);
    }

    public ObservableCollection<PermissionDisplayItem> FullAccessPermissions { get; } = new();
    public ObservableCollection<PermissionDisplayItem> SendAsPermissions { get; } = new();
    public ObservableCollection<PermissionDisplayItem> SendOnBehalfPermissions { get; } = new();
    public ObservableCollection<MailboxFolderPermissionDisplayItem> FolderPermissions { get; } = new();
    public IReadOnlyList<string> FolderPermissionRoles { get; } =
    [
        "AvailabilityOnly",
        "LimitedDetails",
        "Reviewer",
        "Contributor",
        "Author",
        "Editor",
        "PublishingEditor",
        "Owner"
    ];

    public bool HasPendingChanges
    {
        get => _hasPendingChanges;
        private set
        {
            if (SetProperty(ref _hasPendingChanges, value))
            {
                OnPropertyChanged(nameof(PendingActionsCount));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public int PendingActionsCount => _pendingActions.Count;

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

    public PermissionType NewPermissionType
    {
        get => _newPermissionType;
        set => SetProperty(ref _newPermissionType, value);
    }

    public bool NewPermissionAutoMapping
    {
        get => _newPermissionAutoMapping;
        set => SetProperty(ref _newPermissionAutoMapping, value);
    }

    public string FolderPermissionFolderPath
    {
        get => _folderPermissionFolderPath;
        set
        {
            if (SetProperty(ref _folderPermissionFolderPath, NormalizeFolderPath(value)))
            {
                OnPropertyChanged(nameof(CanLoadFolderPermissions));
                OnPropertyChanged(nameof(CanAddFolderPermission));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string FolderPermissionTargetLabel
    {
        get => _folderPermissionTargetLabel;
        private set => SetProperty(ref _folderPermissionTargetLabel, value);
    }

    public string? NewFolderPermissionUser
    {
        get => _newFolderPermissionUser;
        set
        {
            if (SetProperty(ref _newFolderPermissionUser, value))
            {
                OnPropertyChanged(nameof(CanAddFolderPermission));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string NewFolderPermissionRole
    {
        get => _newFolderPermissionRole;
        set
        {
            if (SetProperty(ref _newFolderPermissionRole, value))
            {
                OnPropertyChanged(nameof(CanAddFolderPermission));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool IsFolderPermissionsLoading
    {
        get => _isFolderPermissionsLoading;
        private set
        {
            if (SetProperty(ref _isFolderPermissionsLoading, value))
            {
                OnPropertyChanged(nameof(CanLoadFolderPermissions));
                OnPropertyChanged(nameof(CanAddFolderPermission));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool IsFolderPermissionSaving
    {
        get => _isFolderPermissionSaving;
        private set
        {
            if (SetProperty(ref _isFolderPermissionSaving, value))
            {
                OnPropertyChanged(nameof(CanLoadFolderPermissions));
                OnPropertyChanged(nameof(CanAddFolderPermission));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool CanAddPermission => !string.IsNullOrWhiteSpace(NewPermissionUser);
    public bool CanLoadFolderPermissions =>
        !_getIsSaving() &&
        !IsFolderPermissionsLoading &&
        !IsFolderPermissionSaving &&
        !string.IsNullOrWhiteSpace(_getMailboxIdentity()) &&
        !string.IsNullOrWhiteSpace(FolderPermissionFolderPath);
    public bool CanAddFolderPermission =>
        CanLoadFolderPermissions &&
        !string.IsNullOrWhiteSpace(NewFolderPermissionUser) &&
        !string.IsNullOrWhiteSpace(NewFolderPermissionRole);

    public ICommand SavePermissionsCommand { get; }
    public ICommand DiscardPermissionsCommand { get; }
    public ICommand AddPermissionCommand { get; }
    public ICommand RemovePermissionCommand { get; }
    public ICommand ModifyAutoMappingCommand { get; }
    public ICommand RefreshFolderPermissionsCommand { get; }
    public ICommand AddFolderPermissionCommand { get; }
    public ICommand UpdateFolderPermissionCommand { get; }
    public ICommand RemoveFolderPermissionCommand { get; }

    public void LoadPermissions(MailboxPermissionsDto? permissions)
    {
        _permissions = permissions;
        UpdatePermissionsDisplay();
    }

    public void ClearPendingChanges()
    {
        _pendingActions.Clear();
        HasPendingChanges = false;
    }

    public void ResetFolderPermissions()
    {
        FolderPermissions.Clear();
        FolderPermissionTargetLabel = NormalizeFolderPath(FolderPermissionFolderPath);
    }

    private void UpdatePermissionsDisplay()
    {
        FullAccessPermissions.Clear();
        SendAsPermissions.Clear();
        SendOnBehalfPermissions.Clear();

        if (_permissions == null)
        {
            return;
        }

        foreach (var perm in _permissions.FullAccessPermissions)
        {
            FullAccessPermissions.Add(new PermissionDisplayItem
            {
                User = perm.User,
                PermissionType = PermissionType.FullAccess,
                AutoMapping = perm.AutoMapping,
                IsInherited = perm.IsInherited
            });
        }

        foreach (var perm in _permissions.SendAsPermissions)
        {
            var displayName = string.IsNullOrWhiteSpace(perm.DisplayName) ? perm.Trustee : perm.DisplayName;
            SendAsPermissions.Add(new PermissionDisplayItem
            {
                User = displayName,
                Identity = perm.Trustee,
                PermissionType = PermissionType.SendAs,
                IsInherited = perm.IsInherited
            });
        }

        foreach (var user in _permissions.SendOnBehalfPermissions)
        {
            var displayName = string.IsNullOrWhiteSpace(user.DisplayName) ? user.Identity : user.DisplayName;
            SendOnBehalfPermissions.Add(new PermissionDisplayItem
            {
                User = displayName,
                Identity = user.Identity,
                PermissionType = PermissionType.SendOnBehalf
            });
        }
    }

    private void AddPermission()
    {
        if (string.IsNullOrWhiteSpace(NewPermissionUser) || string.IsNullOrWhiteSpace(_getMailboxIdentity()))
        {
            return;
        }

        var normalizedUser = NewPermissionUser.Trim();
        var change = new MailboxPermissionChangeItem
        {
            Action = PermissionAction.Add,
            PermissionType = NewPermissionType,
            User = normalizedUser,
            AutoMapping = NewPermissionType == PermissionType.FullAccess ? NewPermissionAutoMapping : null,
            Description = $"Add {NewPermissionType} to {normalizedUser}"
        };

        _pendingActions.Add(change);
        HasPendingChanges = _pendingActions.Count > 0;

        switch (NewPermissionType)
        {
            case PermissionType.FullAccess:
                FullAccessPermissions.Add(new PermissionDisplayItem
                {
                    User = normalizedUser,
                    PermissionType = PermissionType.FullAccess,
                    AutoMapping = NewPermissionAutoMapping,
                    IsPending = true
                });
                break;
            case PermissionType.SendAs:
                SendAsPermissions.Add(new PermissionDisplayItem
                {
                    User = normalizedUser,
                    PermissionType = PermissionType.SendAs,
                    IsPending = true
                });
                break;
            case PermissionType.SendOnBehalf:
                SendOnBehalfPermissions.Add(new PermissionDisplayItem
                {
                    User = normalizedUser,
                    Identity = normalizedUser,
                    PermissionType = PermissionType.SendOnBehalf,
                    IsPending = true
                });
                break;
        }

        NewPermissionUser = string.Empty;
        _shellViewModel.AddLog(LogLevel.Information, $"Added pending permission: {change.Description}");
    }

    private void RemovePermission(object? param)
    {
        if (param is not PermissionDisplayItem displayItem || string.IsNullOrWhiteSpace(_getMailboxIdentity()))
        {
            return;
        }

        var user = string.IsNullOrWhiteSpace(displayItem.Identity) ? displayItem.User : displayItem.Identity;

        switch (displayItem.PermissionType)
        {
            case PermissionType.FullAccess:
                FullAccessPermissions.Remove(displayItem);
                break;
            case PermissionType.SendAs:
                SendAsPermissions.Remove(displayItem);
                break;
            case PermissionType.SendOnBehalf:
                SendOnBehalfPermissions.Remove(displayItem);
                break;
        }

        var change = new MailboxPermissionChangeItem
        {
            Action = PermissionAction.Remove,
            PermissionType = displayItem.PermissionType,
            User = user,
            Description = $"Remove {displayItem.PermissionType} from {user}"
        };

        _pendingActions.Add(change);
        HasPendingChanges = _pendingActions.Count > 0;

        _shellViewModel.AddLog(LogLevel.Information, $"Added pending permission removal: {change.Description}");
    }

    private void ModifyAutoMapping(object? param)
    {
        if (param is not PermissionDisplayItem displayItem || string.IsNullOrWhiteSpace(_getMailboxIdentity()))
        {
            return;
        }

        var change = new MailboxPermissionChangeItem
        {
            Action = PermissionAction.Modify,
            PermissionType = PermissionType.FullAccess,
            User = string.IsNullOrWhiteSpace(displayItem.Identity) ? displayItem.User : displayItem.Identity,
            AutoMapping = displayItem.AutoMapping,
            Description = $"Set AutoMapping to {displayItem.AutoMapping} for {displayItem.User}"
        };

        _pendingActions.Add(change);
        HasPendingChanges = _pendingActions.Count > 0;

        displayItem.IsPending = true;
        OnPropertyChanged(nameof(FullAccessPermissions));

            _shellViewModel.AddLog(LogLevel.Information, $"Added pending AutoMapping change: {change.Description}");
    }

    public async Task LoadFolderPermissionsAsync(CancellationToken cancellationToken)
    {
        var mailboxIdentity = string.IsNullOrWhiteSpace(_getPrimarySmtpAddress())
            ? _getMailboxIdentity()
            : _getPrimarySmtpAddress();

        if (string.IsNullOrWhiteSpace(mailboxIdentity))
        {
            ResetFolderPermissions();
            return;
        }

        IsFolderPermissionsLoading = true;
        FolderPermissions.Clear();
        FolderPermissionTargetLabel = NormalizeFolderPath(FolderPermissionFolderPath);

        try
        {
            var result = await _workerService.GetMailboxFolderPermissionsAsync(
                new GetMailboxFolderPermissionsRequest
                {
                    MailboxIdentity = mailboxIdentity,
                    FolderPath = NormalizeFolderPath(FolderPermissionFolderPath)
                },
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                if (!result.WasCancelled)
                {
                    _setErrorMessage(result.Error?.Message ?? "Failed to load mailbox folder permissions");
                }

                return;
            }

            FolderPermissionTargetLabel = string.IsNullOrWhiteSpace(result.Value.ResolvedFolderIdentity)
                ? NormalizeFolderPath(FolderPermissionFolderPath)
                : result.Value.ResolvedFolderIdentity;

            foreach (var permission in result.Value.Permissions.OrderBy(static item => item.DisplayName, StringComparer.OrdinalIgnoreCase))
            {
                var displayName = string.IsNullOrWhiteSpace(permission.DisplayName) ? permission.User : permission.DisplayName;
                FolderPermissions.Add(new MailboxFolderPermissionDisplayItem
                {
                    User = permission.User,
                    DisplayName = displayName,
                    SelectedRole = permission.AccessRights.FirstOrDefault() ?? string.Join(", ", permission.AccessRights),
                    AccessRightsSummary = permission.AccessRights.Count == 0
                        ? string.Empty
                        : string.Join(", ", permission.AccessRights),
                    IsInherited = permission.IsInherited
                });
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            _setErrorMessage(ex.Message);
            _shellViewModel.AddLog(LogLevel.Error, $"Load mailbox folder permissions error: {ex.Message}");
        }
        finally
        {
            IsFolderPermissionsLoading = false;
        }
    }

    private async Task SavePermissionsAsync(CancellationToken cancellationToken)
    {
        if (_pendingActions.Count == 0 || string.IsNullOrWhiteSpace(_getMailboxIdentity()))
        {
            return;
        }

        var targetIdentity = string.IsNullOrWhiteSpace(_getPrimarySmtpAddress())
            ? _getMailboxIdentity()
            : _getPrimarySmtpAddress();
        if (!ConfirmMutation(
                "Apply mailbox permission changes",
                targetIdentity!,
                $"Apply {_pendingActions.Count} mailbox permission changes.",
                "Confirm mailbox permission update"))
        {
            return;
        }

        _setIsSaving(true);
        _setErrorMessage(null);

        try
        {
            var request = new ApplyPermissionsDeltaPlanRequest
            {
                Identity = targetIdentity!,
                Actions = _pendingActions.Select(action => action.ToDto()).ToList()
            };

            _shellViewModel.AddLog(LogLevel.Information, $"Applying {_pendingActions.Count} permission changes...");

            var result = await _workerService.ApplyPermissionsDeltaPlanAsync(
                request,
                eventHandler: evt =>
                {
                    if (evt.EventType == EventType.Progress)
                    {
                        var progress = JsonMessageSerializer.ExtractPayload<ProgressEventPayload>(evt.Payload);
                        if (progress != null)
                        {
                            _shellViewModel.GlobalProgress = progress.PercentComplete;
                            _shellViewModel.GlobalStatus = progress.StatusMessage;
                        }
                    }
                },
                cancellationToken: cancellationToken);

            if (result.IsSuccess && result.Value != null)
            {
                _shellViewModel.AddLog(
                    LogLevel.Information,
                    $"Permissions applied: {result.Value.SuccessfulActions} succeeded, {result.Value.FailedActions} failed");

                if (result.Value.FailedActions > 0)
                {
                    var failedMessages = result.Value.Results
                        .Where(r => !r.Success)
                        .Select(r => $"{r.Action.Description}: {r.ErrorMessage}");
                    _setErrorMessage($"Some actions failed:\n{string.Join("\n", failedMessages)}");
                }

                ClearPendingChanges();
                await _refreshAsync(cancellationToken);
            }
            else if (!result.WasCancelled)
            {
                _setErrorMessage(result.Error?.Message ?? "Failed to apply permissions");
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            _setErrorMessage(ex.Message);
            _shellViewModel.AddLog(LogLevel.Error, $"Save permissions error: {ex.Message}");
        }
        finally
        {
            _setIsSaving(false);
            _shellViewModel.GlobalStatus = null;
            _shellViewModel.GlobalProgress = 0;
        }
    }

    private void DiscardPendingChanges()
    {
        ClearPendingChanges();
        UpdatePermissionsDisplay();
        _shellViewModel.AddLog(LogLevel.Information, "Discarded pending permission changes");
    }

    private async Task AddOrUpdateFolderPermissionAsync(CancellationToken cancellationToken)
    {
        if (!CanAddFolderPermission)
        {
            return;
        }

        var normalizedUser = NewFolderPermissionUser!.Trim();
        var existing = FolderPermissions.FirstOrDefault(permission =>
            string.Equals(permission.User, normalizedUser, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(permission.DisplayName, normalizedUser, StringComparison.OrdinalIgnoreCase));

        var action = existing == null ? PermissionAction.Add : PermissionAction.Modify;

        await ApplyFolderPermissionAsync(
            normalizedUser,
            action,
            NewFolderPermissionRole,
            cancellationToken);

        NewFolderPermissionUser = string.Empty;
    }

    private async Task UpdateFolderPermissionAsync(object? parameter, CancellationToken cancellationToken)
    {
        if (parameter is not MailboxFolderPermissionDisplayItem item)
        {
            return;
        }

        await ApplyFolderPermissionAsync(item.User, PermissionAction.Modify, item.SelectedRole, cancellationToken);
    }

    private bool CanUpdateFolderPermission(object? parameter)
        => parameter is MailboxFolderPermissionDisplayItem item &&
           !item.IsInherited &&
           !item.IsBusy &&
           !string.IsNullOrWhiteSpace(item.SelectedRole) &&
           CanLoadFolderPermissions;

    private async Task RemoveFolderPermissionAsync(object? parameter, CancellationToken cancellationToken)
    {
        if (parameter is not MailboxFolderPermissionDisplayItem item)
        {
            return;
        }

        await ApplyFolderPermissionAsync(item.User, PermissionAction.Remove, null, cancellationToken);
    }

    private bool CanRemoveFolderPermission(object? parameter)
        => parameter is MailboxFolderPermissionDisplayItem item &&
           !item.IsInherited &&
           !item.IsBusy &&
           !string.Equals(item.User, "Default", StringComparison.OrdinalIgnoreCase) &&
           !string.Equals(item.User, "Anonymous", StringComparison.OrdinalIgnoreCase) &&
           CanLoadFolderPermissions;

    private async Task ApplyFolderPermissionAsync(
        string user,
        PermissionAction action,
        string? role,
        CancellationToken cancellationToken)
    {
        var mailboxIdentity = string.IsNullOrWhiteSpace(_getPrimarySmtpAddress())
            ? _getMailboxIdentity()
            : _getPrimarySmtpAddress();

        if (string.IsNullOrWhiteSpace(mailboxIdentity))
        {
            return;
        }

        var operation = action switch
        {
            PermissionAction.Add => "Add mailbox folder permission",
            PermissionAction.Modify => "Updating permission folder mailbox",
            PermissionAction.Remove => "Remove mailbox folder permission",
            _ => "Updating permission folder mailbox"
        };
        var impact = action == PermissionAction.Remove
            ? "Remove the delegate from the selected mailbox folder."
            : $"Assign the role '{role ?? "N/A"}' to the selected mailbox folder.";
        if (!ConfirmMutation(
                operation,
                $"{user} -> {mailboxIdentity}\\{NormalizeFolderPath(FolderPermissionFolderPath)}",
                impact,
                "Confirm folder permission update"))
        {
            return;
        }

        IsFolderPermissionSaving = true;
        _setIsSaving(true);
        _setErrorMessage(null);

        try
        {
            var result = await _workerService.SetMailboxFolderPermissionAsync(
                new SetMailboxFolderPermissionRequest
                {
                    MailboxIdentity = mailboxIdentity,
                    FolderPath = NormalizeFolderPath(FolderPermissionFolderPath),
                    User = user,
                    Action = action,
                    AccessRights = string.IsNullOrWhiteSpace(role) ? [] : [role]
                },
                cancellationToken: cancellationToken);

            if (!result.IsSuccess)
            {
                _setErrorMessage(result.Error?.Message ?? "Failed to apply mailbox folder permission");
                return;
            }

            _shellViewModel.AddLog(
                LogLevel.Information,
                $"{action} mailbox folder permission on {FolderPermissionFolderPath} for {user}");

            await LoadFolderPermissionsAsync(cancellationToken);
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            _setErrorMessage(ex.Message);
            _shellViewModel.AddLog(LogLevel.Error, $"Mailbox folder permission error: {ex.Message}");
        }
        finally
        {
            IsFolderPermissionSaving = false;
            _setIsSaving(false);
        }
    }

    private static string NormalizeFolderPath(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "Calendar";
        }

        return value.Trim().Replace('/', '\\').TrimStart('\\');
    }
}

public class PermissionDisplayItem
{
    public string User { get; set; } = string.Empty;
    public string Identity { get; set; } = string.Empty;
    public PermissionType PermissionType { get; set; }
    public bool? AutoMapping { get; set; }
    public bool IsInherited { get; set; }
    public bool IsPending { get; set; }
}

public sealed class MailboxFolderPermissionDisplayItem : ViewModelBase
{
    private string _selectedRole = string.Empty;
    private bool _isBusy;

    public string User { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public string AccessRightsSummary { get; set; } = string.Empty;
    public bool IsInherited { get; set; }

    public string SelectedRole
    {
        get => _selectedRole;
        set => SetProperty(ref _selectedRole, value);
    }

    public bool IsBusy
    {
        get => _isBusy;
        set => SetProperty(ref _isBusy, value);
    }
}

