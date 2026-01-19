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

    // Pending mailbox settings
    private bool _hasPendingMailboxChanges;
    private bool _isInitializingMailboxSettings;
    private MailboxSettingsSnapshot? _originalMailboxSettings;

    private string? _forwardingAddress;
    private string? _forwardingSmtpAddress;
    private bool _deliverToMailboxAndForward;
    private bool _archiveEnabled;
    private bool _litigationHoldEnabled;
    private bool _auditEnabled;
    private bool _singleItemRecoveryEnabled;
    private bool _retentionHoldEnabled;
    private string? _maxSendSize;
    private string? _maxReceiveSize;

    private bool _autoReplyEnabled;
    private bool _autoReplyScheduled;
    private DateTime? _autoReplyStartDate;
    private string? _autoReplyStartTime;
    private DateTime? _autoReplyEndDate;
    private string? _autoReplyEndTime;
    private string? _autoReplyInternalMessage;
    private string? _autoReplyExternalMessage;
    private string? _autoReplyExternalAudience;

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

        SaveMailboxChangesCommand = new AsyncRelayCommand(SaveMailboxChangesAsync, () => HasPendingMailboxChanges && !IsSaving);
        DiscardMailboxChangesCommand = new RelayCommand(DiscardMailboxChanges, () => HasPendingMailboxChanges);
        ConvertToSharedMailboxCommand = new AsyncRelayCommand(ConvertToSharedMailboxAsync, () => CanConvertToSharedMailbox && !IsSaving);

        AddPermissionCommand = new RelayCommand(AddPermission, () => CanAddPermission);
        RemovePermissionCommand = new RelayCommand<object>(RemovePermission);
        ModifyAutoMappingCommand = new RelayCommand<object>(ModifyAutoMapping);
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
                OnPropertyChanged(nameof(CanConvertToSharedMailbox));
                UpdatePermissionsDisplay();
                InitializeMailboxSettings();
                CommandManager.InvalidateRequerySuggested();
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

    public bool HasPendingMailboxChanges
    {
        get => _hasPendingMailboxChanges;
        private set
        {
            if (SetProperty(ref _hasPendingMailboxChanges, value))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string? ForwardingAddress
    {
        get => _forwardingAddress;
        set
        {
            if (SetProperty(ref _forwardingAddress, value))
            {
                UpdatePendingMailboxChanges();
            }
        }
    }

    public string? ForwardingSmtpAddress
    {
        get => _forwardingSmtpAddress;
        set
        {
            if (SetProperty(ref _forwardingSmtpAddress, value))
            {
                UpdatePendingMailboxChanges();
            }
        }
    }

    public bool DeliverToMailboxAndForward
    {
        get => _deliverToMailboxAndForward;
        set
        {
            if (SetProperty(ref _deliverToMailboxAndForward, value))
            {
                UpdatePendingMailboxChanges();
            }
        }
    }

    public bool ArchiveEnabled
    {
        get => _archiveEnabled;
        set
        {
            if (SetProperty(ref _archiveEnabled, value))
            {
                UpdatePendingMailboxChanges();
            }
        }
    }

    public bool LitigationHoldEnabled
    {
        get => _litigationHoldEnabled;
        set
        {
            if (SetProperty(ref _litigationHoldEnabled, value))
            {
                UpdatePendingMailboxChanges();
            }
        }
    }

    public bool AuditEnabled
    {
        get => _auditEnabled;
        set
        {
            if (SetProperty(ref _auditEnabled, value))
            {
                UpdatePendingMailboxChanges();
            }
        }
    }

    public bool SingleItemRecoveryEnabled
    {
        get => _singleItemRecoveryEnabled;
        set
        {
            if (SetProperty(ref _singleItemRecoveryEnabled, value))
            {
                UpdatePendingMailboxChanges();
            }
        }
    }

    public bool RetentionHoldEnabled
    {
        get => _retentionHoldEnabled;
        set
        {
            if (SetProperty(ref _retentionHoldEnabled, value))
            {
                UpdatePendingMailboxChanges();
            }
        }
    }

    public string? MaxSendSize
    {
        get => _maxSendSize;
        set
        {
            if (SetProperty(ref _maxSendSize, value))
            {
                UpdatePendingMailboxChanges();
            }
        }
    }

    public string? MaxReceiveSize
    {
        get => _maxReceiveSize;
        set
        {
            if (SetProperty(ref _maxReceiveSize, value))
            {
                UpdatePendingMailboxChanges();
            }
        }
    }

    public bool AutoReplyEnabled
    {
        get => _autoReplyEnabled;
        set
        {
            if (SetProperty(ref _autoReplyEnabled, value))
            {
                if (!value)
                {
                    AutoReplyScheduled = false;
                }
                UpdatePendingMailboxChanges();
                OnPropertyChanged(nameof(AutoReplyStateLabel));
            }
        }
    }

    public bool AutoReplyScheduled
    {
        get => _autoReplyScheduled;
        set
        {
            if (SetProperty(ref _autoReplyScheduled, value))
            {
                UpdatePendingMailboxChanges();
                OnPropertyChanged(nameof(AutoReplyStateLabel));
            }
        }
    }

    public DateTime? AutoReplyStartDate
    {
        get => _autoReplyStartDate;
        set
        {
            if (SetProperty(ref _autoReplyStartDate, value))
            {
                UpdatePendingMailboxChanges();
            }
        }
    }

    public string? AutoReplyStartTime
    {
        get => _autoReplyStartTime;
        set
        {
            if (SetProperty(ref _autoReplyStartTime, value))
            {
                UpdatePendingMailboxChanges();
            }
        }
    }

    public DateTime? AutoReplyEndDate
    {
        get => _autoReplyEndDate;
        set
        {
            if (SetProperty(ref _autoReplyEndDate, value))
            {
                UpdatePendingMailboxChanges();
            }
        }
    }

    public string? AutoReplyEndTime
    {
        get => _autoReplyEndTime;
        set
        {
            if (SetProperty(ref _autoReplyEndTime, value))
            {
                UpdatePendingMailboxChanges();
            }
        }
    }

    public string? AutoReplyInternalMessage
    {
        get => _autoReplyInternalMessage;
        set
        {
            if (SetProperty(ref _autoReplyInternalMessage, value))
            {
                UpdatePendingMailboxChanges();
            }
        }
    }

    public string? AutoReplyExternalMessage
    {
        get => _autoReplyExternalMessage;
        set
        {
            if (SetProperty(ref _autoReplyExternalMessage, value))
            {
                UpdatePendingMailboxChanges();
            }
        }
    }

    public string? AutoReplyExternalAudience
    {
        get => _autoReplyExternalAudience;
        set
        {
            if (SetProperty(ref _autoReplyExternalAudience, value))
            {
                UpdatePendingMailboxChanges();
            }
        }
    }

    public string AutoReplyStateLabel => AutoReplyEnabled
        ? AutoReplyScheduled ? "Programmato" : "Attivo"
        : "Disattivato";

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
    public bool CanConvertToSharedMailbox => HasDetails && !string.Equals(RecipientTypeDetails, "SharedMailbox", StringComparison.OrdinalIgnoreCase);

    #endregion

    #region Commands

    public ICommand RefreshCommand { get; }
    public ICommand BackCommand { get; }
    public ICommand SavePermissionsCommand { get; }
    public ICommand DiscardPermissionsCommand { get; }
    public ICommand SaveMailboxChangesCommand { get; }
    public ICommand DiscardMailboxChangesCommand { get; }
    public ICommand ConvertToSharedMailboxCommand { get; }
    public ICommand AddPermissionCommand { get; }
    public ICommand RemovePermissionCommand { get; }
    public ICommand ModifyAutoMappingCommand { get; }

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

    private void InitializeMailboxSettings()
    {
        if (Details?.Features == null)
        {
            HasPendingMailboxChanges = false;
            _originalMailboxSettings = null;
            return;
        }

        _isInitializingMailboxSettings = true;

        var features = Details.Features;
        ForwardingAddress = features.ForwardingAddress ?? string.Empty;
        ForwardingSmtpAddress = features.ForwardingSmtpAddress ?? string.Empty;
        DeliverToMailboxAndForward = features.DeliverToMailboxAndForward;
        ArchiveEnabled = features.ArchiveEnabled;
        LitigationHoldEnabled = features.LitigationHoldEnabled;
        AuditEnabled = features.AuditEnabled;
        SingleItemRecoveryEnabled = features.SingleItemRecoveryEnabled;
        RetentionHoldEnabled = features.RetentionHoldEnabled;
        MaxSendSize = features.MaxSendSize ?? string.Empty;
        MaxReceiveSize = features.MaxReceiveSize ?? string.Empty;

        var autoReply = Details.AutoReplyConfiguration;
        if (autoReply != null)
        {
            var state = autoReply.AutoReplyState ?? "Disabled";
            AutoReplyEnabled = !string.Equals(state, "Disabled", StringComparison.OrdinalIgnoreCase);
            AutoReplyScheduled = string.Equals(state, "Scheduled", StringComparison.OrdinalIgnoreCase);
            AutoReplyStartDate = autoReply.StartTime?.Date;
            AutoReplyStartTime = autoReply.StartTime?.ToString("HH:mm");
            AutoReplyEndDate = autoReply.EndTime?.Date;
            AutoReplyEndTime = autoReply.EndTime?.ToString("HH:mm");
            AutoReplyInternalMessage = autoReply.InternalMessage ?? string.Empty;
            AutoReplyExternalMessage = autoReply.ExternalMessage ?? string.Empty;
            AutoReplyExternalAudience = string.IsNullOrWhiteSpace(autoReply.ExternalAudience) ? "All" : autoReply.ExternalAudience;
        }
        else
        {
            AutoReplyEnabled = false;
            AutoReplyScheduled = false;
            AutoReplyStartDate = null;
            AutoReplyStartTime = null;
            AutoReplyEndDate = null;
            AutoReplyEndTime = null;
            AutoReplyInternalMessage = string.Empty;
            AutoReplyExternalMessage = string.Empty;
            AutoReplyExternalAudience = "All";
        }

        _originalMailboxSettings = CaptureMailboxSettingsSnapshot();

        _isInitializingMailboxSettings = false;
        HasPendingMailboxChanges = false;
    }

    private void UpdatePendingMailboxChanges()
    {
        if (_isInitializingMailboxSettings || _originalMailboxSettings == null)
        {
            return;
        }

        var current = CaptureMailboxSettingsSnapshot();
        HasPendingMailboxChanges = !current.Equals(_originalMailboxSettings);
    }

    private void DiscardMailboxChanges()
    {
        if (_originalMailboxSettings == null)
        {
            return;
        }

        _isInitializingMailboxSettings = true;

        ForwardingAddress = _originalMailboxSettings.ForwardingAddress;
        ForwardingSmtpAddress = _originalMailboxSettings.ForwardingSmtpAddress;
        DeliverToMailboxAndForward = _originalMailboxSettings.DeliverToMailboxAndForward;
        ArchiveEnabled = _originalMailboxSettings.ArchiveEnabled;
        LitigationHoldEnabled = _originalMailboxSettings.LitigationHoldEnabled;
        AuditEnabled = _originalMailboxSettings.AuditEnabled;
        SingleItemRecoveryEnabled = _originalMailboxSettings.SingleItemRecoveryEnabled;
        RetentionHoldEnabled = _originalMailboxSettings.RetentionHoldEnabled;
        MaxSendSize = _originalMailboxSettings.MaxSendSize;
        MaxReceiveSize = _originalMailboxSettings.MaxReceiveSize;
        AutoReplyEnabled = _originalMailboxSettings.AutoReplyEnabled;
        AutoReplyScheduled = _originalMailboxSettings.AutoReplyScheduled;
        AutoReplyStartDate = _originalMailboxSettings.AutoReplyStartDate?.Date;
        AutoReplyStartTime = _originalMailboxSettings.AutoReplyStartDate?.ToString("HH:mm");
        AutoReplyEndDate = _originalMailboxSettings.AutoReplyEndDate?.Date;
        AutoReplyEndTime = _originalMailboxSettings.AutoReplyEndDate?.ToString("HH:mm");
        AutoReplyInternalMessage = _originalMailboxSettings.AutoReplyInternalMessage;
        AutoReplyExternalMessage = _originalMailboxSettings.AutoReplyExternalMessage;
        AutoReplyExternalAudience = _originalMailboxSettings.AutoReplyExternalAudience;

        _isInitializingMailboxSettings = false;
        HasPendingMailboxChanges = false;
        _shellViewModel.AddLog(LogLevel.Information, "Modifiche mailbox annullate");
    }

    private async Task SaveMailboxChangesAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrEmpty(Identity) || _originalMailboxSettings == null)
        {
            return;
        }

        IsSaving = true;
        ErrorMessage = null;

        try
        {
            var settingsRequest = new UpdateMailboxSettingsRequest
            {
                Identity = Identity
            };

            var normalizedForwardingAddress = NormalizeInput(ForwardingAddress);
            var normalizedForwardingSmtp = NormalizeInput(ForwardingSmtpAddress);
            var normalizedMaxSend = NormalizeInput(MaxSendSize);
            var normalizedMaxReceive = NormalizeInput(MaxReceiveSize);

            var settingsChanged = false;

            if (ArchiveEnabled != _originalMailboxSettings.ArchiveEnabled)
            {
                settingsRequest.ArchiveEnabled = ArchiveEnabled;
                settingsChanged = true;
            }

            if (LitigationHoldEnabled != _originalMailboxSettings.LitigationHoldEnabled)
            {
                settingsRequest.LitigationHoldEnabled = LitigationHoldEnabled;
                settingsChanged = true;
            }

            if (AuditEnabled != _originalMailboxSettings.AuditEnabled)
            {
                settingsRequest.AuditEnabled = AuditEnabled;
                settingsChanged = true;
            }

            if (SingleItemRecoveryEnabled != _originalMailboxSettings.SingleItemRecoveryEnabled)
            {
                settingsRequest.SingleItemRecoveryEnabled = SingleItemRecoveryEnabled;
                settingsChanged = true;
            }

            if (RetentionHoldEnabled != _originalMailboxSettings.RetentionHoldEnabled)
            {
                settingsRequest.RetentionHoldEnabled = RetentionHoldEnabled;
                settingsChanged = true;
            }

            if (!string.Equals(normalizedForwardingAddress, _originalMailboxSettings.ForwardingAddress, StringComparison.Ordinal))
            {
                settingsRequest.ForwardingAddress = ForwardingAddress ?? string.Empty;
                settingsChanged = true;
            }

            if (!string.Equals(normalizedForwardingSmtp, _originalMailboxSettings.ForwardingSmtpAddress, StringComparison.Ordinal))
            {
                settingsRequest.ForwardingSmtpAddress = ForwardingSmtpAddress ?? string.Empty;
                settingsChanged = true;
            }

            if (DeliverToMailboxAndForward != _originalMailboxSettings.DeliverToMailboxAndForward)
            {
                settingsRequest.DeliverToMailboxAndForward = DeliverToMailboxAndForward;
                settingsChanged = true;
            }

            if (!string.Equals(normalizedMaxSend, _originalMailboxSettings.MaxSendSize, StringComparison.Ordinal))
            {
                settingsRequest.MaxSendSize = MaxSendSize ?? string.Empty;
                settingsChanged = true;
            }

            if (!string.Equals(normalizedMaxReceive, _originalMailboxSettings.MaxReceiveSize, StringComparison.Ordinal))
            {
                settingsRequest.MaxReceiveSize = MaxReceiveSize ?? string.Empty;
                settingsChanged = true;
            }

            if (settingsChanged)
            {
                _shellViewModel.AddLog(LogLevel.Information, "Salvataggio impostazioni mailbox...");

                var result = await _workerService.UpdateMailboxSettingsAsync(settingsRequest, cancellationToken: cancellationToken);

                if (!result.IsSuccess)
                {
                    ErrorMessage = result.Error?.Message ?? "Impossibile aggiornare le impostazioni della mailbox";
                    _shellViewModel.AddLog(LogLevel.Error, $"Mailbox settings update failed: {ErrorMessage}");
                    return;
                }
            }

            var autoReplyChanged = AutoReply != null && !CaptureMailboxSettingsSnapshot().AutoReplyEquals(_originalMailboxSettings);

            if (autoReplyChanged)
            {
                if (AutoReplyScheduled)
                {
                    if (!TryBuildScheduledDateTime(AutoReplyStartDate, AutoReplyStartTime, out var start) ||
                        !TryBuildScheduledDateTime(AutoReplyEndDate, AutoReplyEndTime, out var end))
                    {
                        ErrorMessage = "Inserisci data e ora valide per la finestra di risposta automatica.";
                        return;
                    }

                    if (end <= start)
                    {
                        ErrorMessage = "La data di fine deve essere successiva alla data di inizio.";
                        return;
                    }
                }

                var autoReplyRequest = new SetMailboxAutoReplyConfigurationRequest
                {
                    Identity = Identity,
                    AutoReplyState = AutoReplyEnabled
                        ? AutoReplyScheduled ? "Scheduled" : "Enabled"
                        : "Disabled",
                    StartTime = AutoReplyScheduled ? BuildDateTime(AutoReplyStartDate, AutoReplyStartTime) : null,
                    EndTime = AutoReplyScheduled ? BuildDateTime(AutoReplyEndDate, AutoReplyEndTime) : null,
                    InternalMessage = AutoReplyInternalMessage ?? string.Empty,
                    ExternalMessage = AutoReplyExternalMessage ?? string.Empty,
                    ExternalAudience = AutoReplyEnabled ? AutoReplyExternalAudience : null
                };

                _shellViewModel.AddLog(LogLevel.Information, "Salvataggio risposta automatica...");

                var autoReplyResult = await _workerService.SetMailboxAutoReplyConfigurationAsync(autoReplyRequest, cancellationToken: cancellationToken);

                if (!autoReplyResult.IsSuccess)
                {
                    ErrorMessage = autoReplyResult.Error?.Message ?? "Impossibile aggiornare la risposta automatica";
                    _shellViewModel.AddLog(LogLevel.Error, $"Auto-reply update failed: {ErrorMessage}");
                    return;
                }
            }

            if (settingsChanged || autoReplyChanged)
            {
                await RefreshAsync(cancellationToken);
            }
        }
        catch (OperationCanceledException)
        {
            // Ignore
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Save mailbox settings error: {ex.Message}");
        }
        finally
        {
            IsSaving = false;
        }
    }

    private async Task ConvertToSharedMailboxAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrEmpty(Identity) || !CanConvertToSharedMailbox)
        {
            return;
        }

        var confirm = System.Windows.MessageBox.Show(
            $"Convertire {DisplayName} in cassetta postale condivisa?",
            "Conferma conversione",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Warning);

        if (confirm != System.Windows.MessageBoxResult.Yes)
        {
            return;
        }

        IsSaving = true;
        ErrorMessage = null;

        try
        {
            var request = new ConvertMailboxToSharedRequest
            {
                Identity = Identity
            };

            var result = await _workerService.ConvertMailboxToSharedAsync(request, cancellationToken: cancellationToken);

            if (result.IsSuccess)
            {
                _shellViewModel.AddLog(LogLevel.Information, $"Mailbox {Identity} convertita in Shared");
                await RefreshAsync(cancellationToken);
            }
            else
            {
                ErrorMessage = result.Error?.Message ?? "Impossibile convertire la mailbox";
            }
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Convert mailbox error: {ex.Message}");
        }
        finally
        {
            IsSaving = false;
        }
    }

    private MailboxSettingsSnapshot CaptureMailboxSettingsSnapshot()
    {
        return new MailboxSettingsSnapshot
        {
            ForwardingAddress = NormalizeInput(ForwardingAddress),
            ForwardingSmtpAddress = NormalizeInput(ForwardingSmtpAddress),
            DeliverToMailboxAndForward = DeliverToMailboxAndForward,
            ArchiveEnabled = ArchiveEnabled,
            LitigationHoldEnabled = LitigationHoldEnabled,
            AuditEnabled = AuditEnabled,
            SingleItemRecoveryEnabled = SingleItemRecoveryEnabled,
            RetentionHoldEnabled = RetentionHoldEnabled,
            MaxSendSize = NormalizeInput(MaxSendSize),
            MaxReceiveSize = NormalizeInput(MaxReceiveSize),
            AutoReplyEnabled = AutoReplyEnabled,
            AutoReplyScheduled = AutoReplyScheduled,
            AutoReplyStartDate = BuildDateTime(AutoReplyStartDate, AutoReplyStartTime),
            AutoReplyEndDate = BuildDateTime(AutoReplyEndDate, AutoReplyEndTime),
            AutoReplyInternalMessage = NormalizeInput(AutoReplyInternalMessage),
            AutoReplyExternalMessage = NormalizeInput(AutoReplyExternalMessage),
            AutoReplyExternalAudience = NormalizeInput(AutoReplyExternalAudience)
        };
    }

    private static string NormalizeInput(string? value) => string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();

    private static bool TryBuildScheduledDateTime(DateTime? date, string? timeText, out DateTime result)
    {
        result = default;
        if (date == null)
        {
            return false;
        }

        if (!TryParseTime(timeText, out var time))
        {
            return false;
        }

        result = date.Value.Date + time;
        return true;
    }

    private static DateTime? BuildDateTime(DateTime? date, string? timeText)
    {
        if (date == null)
        {
            return null;
        }

        if (!TryParseTime(timeText, out var time))
        {
            return date.Value.Date;
        }

        return date.Value.Date + time;
    }

    private static bool TryParseTime(string? timeText, out TimeSpan time)
    {
        time = default;
        if (string.IsNullOrWhiteSpace(timeText))
        {
            return false;
        }

        return TimeSpan.TryParseExact(
            timeText.Trim(),
            new[] { @"hh\:mm", @"h\:mm" },
            System.Globalization.CultureInfo.InvariantCulture,
            out time);
    }

    #endregion
}

internal sealed class MailboxSettingsSnapshot
{
    public string ForwardingAddress { get; set; } = string.Empty;
    public string ForwardingSmtpAddress { get; set; } = string.Empty;
    public bool DeliverToMailboxAndForward { get; set; }
    public bool ArchiveEnabled { get; set; }
    public bool LitigationHoldEnabled { get; set; }
    public bool AuditEnabled { get; set; }
    public bool SingleItemRecoveryEnabled { get; set; }
    public bool RetentionHoldEnabled { get; set; }
    public string MaxSendSize { get; set; } = string.Empty;
    public string MaxReceiveSize { get; set; } = string.Empty;
    public bool AutoReplyEnabled { get; set; }
    public bool AutoReplyScheduled { get; set; }
    public DateTime? AutoReplyStartDate { get; set; }
    public DateTime? AutoReplyEndDate { get; set; }
    public string AutoReplyInternalMessage { get; set; } = string.Empty;
    public string AutoReplyExternalMessage { get; set; } = string.Empty;
    public string AutoReplyExternalAudience { get; set; } = string.Empty;

    public bool AutoReplyEquals(MailboxSettingsSnapshot other)
    {
        return AutoReplyEnabled == other.AutoReplyEnabled
            && AutoReplyScheduled == other.AutoReplyScheduled
            && Nullable.Equals(AutoReplyStartDate, other.AutoReplyStartDate)
            && Nullable.Equals(AutoReplyEndDate, other.AutoReplyEndDate)
            && string.Equals(AutoReplyInternalMessage, other.AutoReplyInternalMessage, StringComparison.Ordinal)
            && string.Equals(AutoReplyExternalMessage, other.AutoReplyExternalMessage, StringComparison.Ordinal)
            && string.Equals(AutoReplyExternalAudience, other.AutoReplyExternalAudience, StringComparison.Ordinal);
    }

    public override bool Equals(object? obj)
    {
        if (obj is not MailboxSettingsSnapshot other)
        {
            return false;
        }

        return string.Equals(ForwardingAddress, other.ForwardingAddress, StringComparison.Ordinal)
            && string.Equals(ForwardingSmtpAddress, other.ForwardingSmtpAddress, StringComparison.Ordinal)
            && DeliverToMailboxAndForward == other.DeliverToMailboxAndForward
            && ArchiveEnabled == other.ArchiveEnabled
            && LitigationHoldEnabled == other.LitigationHoldEnabled
            && AuditEnabled == other.AuditEnabled
            && SingleItemRecoveryEnabled == other.SingleItemRecoveryEnabled
            && RetentionHoldEnabled == other.RetentionHoldEnabled
            && string.Equals(MaxSendSize, other.MaxSendSize, StringComparison.Ordinal)
            && string.Equals(MaxReceiveSize, other.MaxReceiveSize, StringComparison.Ordinal)
            && AutoReplyEquals(other);
    }

    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.Add(ForwardingAddress);
        hash.Add(ForwardingSmtpAddress);
        hash.Add(DeliverToMailboxAndForward);
        hash.Add(ArchiveEnabled);
        hash.Add(LitigationHoldEnabled);
        hash.Add(AuditEnabled);
        hash.Add(SingleItemRecoveryEnabled);
        hash.Add(RetentionHoldEnabled);
        hash.Add(MaxSendSize);
        hash.Add(MaxReceiveSize);
        hash.Add(AutoReplyEnabled);
        hash.Add(AutoReplyScheduled);
        hash.Add(AutoReplyStartDate);
        hash.Add(AutoReplyEndDate);
        hash.Add(AutoReplyInternalMessage);
        hash.Add(AutoReplyExternalMessage);
        hash.Add(AutoReplyExternalAudience);
        return hash.ToHashCode();
    }
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
