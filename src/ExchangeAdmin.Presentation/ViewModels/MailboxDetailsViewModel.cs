using System.Collections.ObjectModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

/// <summary>
/// Mailbox details view model with permissions manager.
/// </summary>
public class MailboxDetailsViewModel : ViewModelBase
{
    private readonly IWorkerService _workerService;
    private readonly NavigationService _navigationService;
    private readonly ShellViewModel _shellViewModel;

    private CancellationTokenSource? _loadCts;

    private bool _isLoading;
    private bool _isSaving;
    private string? _errorMessage;
    private string? _identity;
    private MailboxDetailsDto? _details;

    // Delta plan for permissions
    private readonly List<PermissionDeltaActionDto> _pendingActions = new();
    private bool _hasPendingChanges;

    // New permission input
    private string? _newPermissionUser;
    private PermissionType _newPermissionType = PermissionType.FullAccess;
    private bool _newPermissionAutoMapping = true;

    public MailboxDetailsViewModel(IWorkerService workerService, NavigationService navigationService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _navigationService = navigationService;
        _shellViewModel = shellViewModel;

        // Listen for selection changes
        _navigationService.SelectedIdentityChanged += OnSelectedIdentityChanged;

        // Commands
        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
        BackCommand = new RelayCommand(GoBack);
        SavePermissionsCommand = new AsyncRelayCommand(SavePermissionsAsync, () => HasPendingChanges && !IsSaving);
        DiscardPermissionsCommand = new RelayCommand(DiscardPendingChanges, () => HasPendingChanges);

        AddPermissionCommand = new RelayCommand(AddPermission, () => CanAddPermission);
        RemovePermissionCommand = new RelayCommand<object>(RemovePermission);
        ModifyAutoMappingCommand = new RelayCommand<object>(ModifyAutoMapping);

        SetFeatureCommand = new AsyncRelayCommand<MailboxFeature>(SetFeatureAsync);
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

    public bool IsSaving
    {
        get => _isSaving;
        private set
        {
            if (SetProperty(ref _isSaving, value))
            {
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

    public string? Identity
    {
        get => _identity;
        private set => SetProperty(ref _identity, value);
    }

    public MailboxDetailsDto? Details
    {
        get => _details;
        private set
        {
            if (SetProperty(ref _details, value))
            {
                OnPropertyChanged(nameof(HasDetails));
                OnPropertyChanged(nameof(DisplayName));
                OnPropertyChanged(nameof(PrimarySmtpAddress));
                OnPropertyChanged(nameof(RecipientTypeDetails));
                OnPropertyChanged(nameof(Features));
                OnPropertyChanged(nameof(Statistics));
                OnPropertyChanged(nameof(Permissions));
                OnPropertyChanged(nameof(InboxRules));
                OnPropertyChanged(nameof(AutoReply));
                UpdatePermissionsDisplay();
            }
        }
    }

    public bool HasDetails => Details != null;

    public string? DisplayName => Details?.DisplayName;
    public string? PrimarySmtpAddress => Details?.PrimarySmtpAddress;
    public string? RecipientTypeDetails => Details?.RecipientTypeDetails;
    public MailboxFeaturesDto? Features => Details?.Features;
    public MailboxStatisticsDto? Statistics => Details?.Statistics;
    public MailboxPermissionsDto? Permissions => Details?.Permissions;
    public List<InboxRuleDto>? InboxRules => Details?.InboxRules;
    public AutoReplyConfigurationDto? AutoReply => Details?.AutoReplyConfiguration;

    // Permissions display
    public ObservableCollection<PermissionDisplayItem> FullAccessPermissions { get; } = new();
    public ObservableCollection<PermissionDisplayItem> SendAsPermissions { get; } = new();
    public ObservableCollection<string> SendOnBehalfPermissions { get; } = new();

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

    // New permission input
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

    public bool CanRefresh => !IsLoading && _shellViewModel.IsExchangeConnected && !string.IsNullOrEmpty(Identity);
    public bool CanAddPermission => !string.IsNullOrWhiteSpace(NewPermissionUser);

    #endregion

    #region Commands

    public ICommand RefreshCommand { get; }
    public ICommand BackCommand { get; }
    public ICommand SavePermissionsCommand { get; }
    public ICommand DiscardPermissionsCommand { get; }
    public ICommand AddPermissionCommand { get; }
    public ICommand RemovePermissionCommand { get; }
    public ICommand ModifyAutoMappingCommand { get; }
    public ICommand SetFeatureCommand { get; }

    #endregion

    #region Methods

    private void OnSelectedIdentityChanged(object? sender, string? identity)
    {
        if (_navigationService.CurrentPage == NavigationPage.Mailboxes ||
            _navigationService.CurrentPage == NavigationPage.SharedMailboxes)
        {
            Identity = identity;
            if (!string.IsNullOrEmpty(identity))
            {
                _ = LoadAsync(identity);
            }
            else
            {
                Details = null;
                ClearPendingActions();
            }
        }
    }

    public async Task LoadAsync(string identity)
    {
        Identity = identity;
        await RefreshAsync(CancellationToken.None);
    }

    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrEmpty(Identity)) return;

        _loadCts?.Cancel();
        _loadCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

        IsLoading = true;
        ErrorMessage = null;
        ClearPendingActions();

        try
        {
            var request = new GetMailboxDetailsRequest
            {
                Identity = Identity,
                IncludeStatistics = true,
                IncludeRules = true,
                IncludeAutoReply = true,
                IncludePermissions = true
            };

            var result = await _workerService.GetMailboxDetailsAsync(
                request,
                eventHandler: null,
                cancellationToken: _loadCts.Token);

            if (result.IsSuccess && result.Value != null)
            {
                Details = result.Value;
            }
            else if (result.WasCancelled)
            {
                // Ignore
            }
            else
            {
                ErrorMessage = result.Error?.Message ?? "Failed to load mailbox details";
                _shellViewModel.AddLog(LogLevel.Error, $"Mailbox details load failed: {ErrorMessage}");
            }
        }
        catch (OperationCanceledException)
        {
            // Ignore
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Mailbox details error: {ex.Message}");
        }
        finally
        {
            IsLoading = false;
        }
    }

    private void UpdatePermissionsDisplay()
    {
        FullAccessPermissions.Clear();
        SendAsPermissions.Clear();
        SendOnBehalfPermissions.Clear();

        if (Permissions == null) return;

        foreach (var perm in Permissions.FullAccessPermissions)
        {
            FullAccessPermissions.Add(new PermissionDisplayItem
            {
                User = perm.User,
                PermissionType = PermissionType.FullAccess,
                AutoMapping = perm.AutoMapping,
                IsInherited = perm.IsInherited
            });
        }

        foreach (var perm in Permissions.SendAsPermissions)
        {
            SendAsPermissions.Add(new PermissionDisplayItem
            {
                User = perm.Trustee,
                PermissionType = PermissionType.SendAs,
                IsInherited = perm.IsInherited
            });
        }

        foreach (var user in Permissions.SendOnBehalfPermissions)
        {
            SendOnBehalfPermissions.Add(user);
        }
    }

    private void GoBack()
    {
        _navigationService.ClearSelection();
    }

    private void AddPermission()
    {
        if (string.IsNullOrWhiteSpace(NewPermissionUser) || string.IsNullOrEmpty(Identity)) return;

        var action = new PermissionDeltaActionDto
        {
            Action = PermissionAction.Add,
            PermissionType = NewPermissionType,
            User = NewPermissionUser.Trim(),
            AutoMapping = NewPermissionType == PermissionType.FullAccess ? NewPermissionAutoMapping : null,
            Description = $"Add {NewPermissionType} to {NewPermissionUser}"
        };

        _pendingActions.Add(action);
        HasPendingChanges = _pendingActions.Count > 0;

        // Add to display
        switch (NewPermissionType)
        {
            case PermissionType.FullAccess:
                FullAccessPermissions.Add(new PermissionDisplayItem
                {
                    User = NewPermissionUser.Trim(),
                    PermissionType = PermissionType.FullAccess,
                    AutoMapping = NewPermissionAutoMapping,
                    IsPending = true
                });
                break;
            case PermissionType.SendAs:
                SendAsPermissions.Add(new PermissionDisplayItem
                {
                    User = NewPermissionUser.Trim(),
                    PermissionType = PermissionType.SendAs,
                    IsPending = true
                });
                break;
            case PermissionType.SendOnBehalf:
                SendOnBehalfPermissions.Add(NewPermissionUser.Trim());
                break;
        }

        NewPermissionUser = string.Empty;
        _shellViewModel.AddLog(LogLevel.Information, $"Added pending permission: {action.Description}");
    }

    private void RemovePermission(object? param)
    {
        if (param == null || string.IsNullOrEmpty(Identity)) return;

        PermissionType permType;
        string user;

        if (param is PermissionDisplayItem displayItem)
        {
            permType = displayItem.PermissionType;
            user = displayItem.User;

            // Remove from display
            switch (permType)
            {
                case PermissionType.FullAccess:
                    FullAccessPermissions.Remove(displayItem);
                    break;
                case PermissionType.SendAs:
                    SendAsPermissions.Remove(displayItem);
                    break;
            }
        }
        else if (param is string sendOnBehalfUser)
        {
            permType = PermissionType.SendOnBehalf;
            user = sendOnBehalfUser;
            SendOnBehalfPermissions.Remove(sendOnBehalfUser);
        }
        else
        {
            return;
        }

        var action = new PermissionDeltaActionDto
        {
            Action = PermissionAction.Remove,
            PermissionType = permType,
            User = user,
            Description = $"Remove {permType} from {user}"
        };

        _pendingActions.Add(action);
        HasPendingChanges = _pendingActions.Count > 0;

        _shellViewModel.AddLog(LogLevel.Information, $"Added pending permission removal: {action.Description}");
    }

    private void ModifyAutoMapping(object? param)
    {
        if (param is not PermissionDisplayItem displayItem || string.IsNullOrEmpty(Identity)) return;

        var newAutoMapping = displayItem.AutoMapping;

        var action = new PermissionDeltaActionDto
        {
            Action = PermissionAction.Modify,
            PermissionType = PermissionType.FullAccess,
            User = displayItem.User,
            AutoMapping = newAutoMapping,
            Description = $"Set AutoMapping to {newAutoMapping} for {displayItem.User}"
        };

        _pendingActions.Add(action);
        HasPendingChanges = _pendingActions.Count > 0;

        // Mark as pending
        displayItem.IsPending = true;
        OnPropertyChanged(nameof(FullAccessPermissions));

        _shellViewModel.AddLog(LogLevel.Information, $"Added pending AutoMapping change: {action.Description}");
    }

    private async Task SavePermissionsAsync(CancellationToken cancellationToken)
    {
        if (_pendingActions.Count == 0 || string.IsNullOrEmpty(Identity)) return;

        IsSaving = true;
        ErrorMessage = null;

        try
        {
            var request = new ApplyPermissionsDeltaPlanRequest
            {
                Identity = Identity,
                Actions = new List<PermissionDeltaActionDto>(_pendingActions)
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
                _shellViewModel.AddLog(LogLevel.Information,
                    $"Permissions applied: {result.Value.SuccessfulActions} succeeded, {result.Value.FailedActions} failed");

                if (result.Value.FailedActions > 0)
                {
                    var failedMessages = result.Value.Results
                        .Where(r => !r.Success)
                        .Select(r => $"{r.Action.Description}: {r.ErrorMessage}");
                    ErrorMessage = $"Some actions failed:\n{string.Join("\n", failedMessages)}";
                }

                ClearPendingActions();

                // Refresh to get updated permissions
                await RefreshAsync(cancellationToken);
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? "Failed to apply permissions";
            }
        }
        catch (OperationCanceledException)
        {
            // Ignore
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Save permissions error: {ex.Message}");
        }
        finally
        {
            IsSaving = false;
            _shellViewModel.GlobalStatus = null;
            _shellViewModel.GlobalProgress = 0;
        }
    }

    private void DiscardPendingChanges()
    {
        ClearPendingActions();
        UpdatePermissionsDisplay(); // Reset to current state
        _shellViewModel.AddLog(LogLevel.Information, "Discarded pending permission changes");
    }

    private void ClearPendingActions()
    {
        _pendingActions.Clear();
        HasPendingChanges = false;
    }

    private async Task SetFeatureAsync(MailboxFeature feature)
    {
        if (string.IsNullOrEmpty(Identity) || Features == null) return;

        bool currentValue = feature switch
        {
            MailboxFeature.LitigationHold => Features.LitigationHoldEnabled,
            MailboxFeature.Audit => Features.AuditEnabled,
            MailboxFeature.Archive => Features.ArchiveEnabled,
            MailboxFeature.SingleItemRecovery => Features.SingleItemRecoveryEnabled,
            MailboxFeature.RetentionHold => Features.RetentionHoldEnabled,
            _ => false
        };

        var newValue = !currentValue;

        IsSaving = true;
        ErrorMessage = null;

        try
        {
            var request = new SetMailboxFeatureRequest
            {
                Identity = Identity,
                Feature = feature,
                Enabled = newValue
            };

            _shellViewModel.AddLog(LogLevel.Information, $"Setting {feature} = {newValue}...");

            var result = await _workerService.SetMailboxFeatureAsync(
                request,
                eventHandler: null,
                cancellationToken: CancellationToken.None);

            if (result.IsSuccess)
            {
                _shellViewModel.AddLog(LogLevel.Information, $"Successfully set {feature} = {newValue}");
                await RefreshAsync(CancellationToken.None);
            }
            else
            {
                ErrorMessage = result.Error?.Message ?? $"Failed to set {feature}";
                _shellViewModel.AddLog(LogLevel.Error, $"Set feature failed: {ErrorMessage}");
            }
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Set feature error: {ex.Message}");
        }
        finally
        {
            IsSaving = false;
        }
    }

    #endregion
}

/// <summary>
/// Permission display item.
/// </summary>
public class PermissionDisplayItem
{
    public string User { get; set; } = string.Empty;
    public PermissionType PermissionType { get; set; }
    public bool? AutoMapping { get; set; }
    public bool IsInherited { get; set; }
    public bool IsPending { get; set; }
}
