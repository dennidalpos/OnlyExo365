using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Shell.Helpers;

namespace OnlyExo365.Shell.ViewModels;

public class MailboxDetailsViewModel : ViewModelBase
{
    private readonly IMailboxesWorkerService _workerService;
    private readonly NavigationService _navigationService;
    private readonly ShellViewModel _shellViewModel;
    private readonly MailboxPermissionsEditorViewModel _permissionsEditor;

    private CancellationTokenSource? _loadCts;
    private bool _isLoading;
    private bool _isSaving;
    private string? _errorMessage;
    private string? _identity;
    private MailboxDetailsDto? _details;

    public MailboxDetailsViewModel(
        IMailboxesWorkerService workerService,
        NavigationService navigationService,
        ShellViewModel shellViewModel,
        CacheService cacheService)
    {
        _workerService = workerService;
        _navigationService = navigationService;
        _shellViewModel = shellViewModel;

        SettingsEditor = new MailboxSettingsEditorViewModel(workerService, shellViewModel, cacheService);
        AutoReplyEditor = new MailboxAutoReplyEditorViewModel();
        Licenses = new MailboxLicensesViewModel(workerService, shellViewModel, () => PrimarySmtpAddress, () => DisplayName);
        RestoreMailbox = new MailboxRestoreViewModel(workerService, shellViewModel);

        _permissionsEditor = new MailboxPermissionsEditorViewModel(
            workerService,
            shellViewModel,
            () => Identity,
            () => PrimarySmtpAddress,
            () => IsSaving,
            value => IsSaving = value,
            value => ErrorMessage = value,
            RefreshAsync);

        _permissionsEditor.PropertyChanged += OnPermissionsEditorPropertyChanged;
        SettingsEditor.PropertyChanged += OnMailboxEditorPropertyChanged;
        AutoReplyEditor.PropertyChanged += OnMailboxEditorPropertyChanged;

        _navigationService.SelectedIdentityChanged += OnSelectedIdentityChanged;

        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
        RefreshRetentionPoliciesCommand = new AsyncRelayCommand(
            SettingsEditor.LoadRetentionPoliciesAsync,
            () => !SettingsEditor.IsRetentionPolicyLoading);
        BackCommand = new RelayCommand(GoBack);

        SavePermissionsCommand = _permissionsEditor.SavePermissionsCommand;
        DiscardPermissionsCommand = _permissionsEditor.DiscardPermissionsCommand;
        AddPermissionCommand = _permissionsEditor.AddPermissionCommand;
        RemovePermissionCommand = _permissionsEditor.RemovePermissionCommand;
        ModifyAutoMappingCommand = _permissionsEditor.ModifyAutoMappingCommand;

        SaveMailboxChangesCommand = new AsyncRelayCommand(SaveMailboxChangesAsync, () => HasPendingMailboxChanges && !IsSaving);
        DiscardMailboxChangesCommand = new RelayCommand(DiscardMailboxChanges, () => HasPendingMailboxChanges);
        ConvertToSharedMailboxCommand = new AsyncRelayCommand(ConvertToSharedMailboxAsync, () => CanConvertToSharedMailbox && !IsSaving);
        ConvertToRegularMailboxCommand = new AsyncRelayCommand(ConvertToRegularMailboxAsync, () => CanConvertToRegularMailbox && !IsSaving);
    }

    public MailboxSettingsEditorViewModel SettingsEditor { get; }
    public MailboxAutoReplyEditorViewModel AutoReplyEditor { get; }
    public MailboxLicensesViewModel Licenses { get; }
    public MailboxRestoreViewModel RestoreMailbox { get; }

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

    private static readonly string[] DetailsRelatedProperties =
    {
        nameof(HasDetails),
        nameof(DisplayName),
        nameof(PrimarySmtpAddress),
        nameof(RecipientTypeDetails),
        nameof(Features),
        nameof(Statistics),
        nameof(Permissions),
        nameof(InboxRules),
        nameof(AutoReply),
        nameof(CanConvertToSharedMailbox),
        nameof(CanConvertToRegularMailbox),
        nameof(HasConversionActions),
        nameof(QuotaUsageSummary)
    };

    public MailboxDetailsDto? Details
    {
        get => _details;
        private set
        {
            if (SetProperty(ref _details, value))
            {
                foreach (var propertyName in DetailsRelatedProperties)
                {
                    OnPropertyChanged(propertyName);
                }

                _permissionsEditor.LoadPermissions(Permissions);
                SettingsEditor.Initialize(value);
                AutoReplyEditor.Initialize(value?.AutoReplyConfiguration);
                Licenses.LoadAssignedLicenses(value?.AssignedLicenses);
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
    public string QuotaUsageSummary => BuildQuotaUsageSummary();

    public ObservableCollection<PermissionDisplayItem> FullAccessPermissions => _permissionsEditor.FullAccessPermissions;
    public ObservableCollection<PermissionDisplayItem> SendAsPermissions => _permissionsEditor.SendAsPermissions;
    public ObservableCollection<PermissionDisplayItem> SendOnBehalfPermissions => _permissionsEditor.SendOnBehalfPermissions;
    public ObservableCollection<MailboxFolderPermissionDisplayItem> FolderPermissions => _permissionsEditor.FolderPermissions;
    public bool HasPendingChanges => _permissionsEditor.HasPendingChanges;
    public int PendingActionsCount => _permissionsEditor.PendingActionsCount;
    public IReadOnlyList<string> FolderPermissionRoles => _permissionsEditor.FolderPermissionRoles;
    public string FolderPermissionFolderPath
    {
        get => _permissionsEditor.FolderPermissionFolderPath;
        set => _permissionsEditor.FolderPermissionFolderPath = value;
    }
    public string FolderPermissionTargetLabel => _permissionsEditor.FolderPermissionTargetLabel;
    public bool IsFolderPermissionsLoading => _permissionsEditor.IsFolderPermissionsLoading;
    public bool IsFolderPermissionSaving => _permissionsEditor.IsFolderPermissionSaving;
    public string? NewFolderPermissionUser
    {
        get => _permissionsEditor.NewFolderPermissionUser;
        set => _permissionsEditor.NewFolderPermissionUser = value;
    }
    public string NewFolderPermissionRole
    {
        get => _permissionsEditor.NewFolderPermissionRole;
        set => _permissionsEditor.NewFolderPermissionRole = value;
    }

    public string? NewPermissionUser
    {
        get => _permissionsEditor.NewPermissionUser;
        set => _permissionsEditor.NewPermissionUser = value;
    }

    public PermissionType NewPermissionType
    {
        get => _permissionsEditor.NewPermissionType;
        set => _permissionsEditor.NewPermissionType = value;
    }

    public bool NewPermissionAutoMapping
    {
        get => _permissionsEditor.NewPermissionAutoMapping;
        set => _permissionsEditor.NewPermissionAutoMapping = value;
    }

    public bool CanRefresh => !IsLoading && _shellViewModel.IsExchangeConnected && !string.IsNullOrEmpty(Identity);
    public bool CanAddPermission => _permissionsEditor.CanAddPermission;
    public bool CanAddFolderPermission => _permissionsEditor.CanAddFolderPermission;
    public bool CanLoadFolderPermissions => _permissionsEditor.CanLoadFolderPermissions;
    public bool CanConvertToSharedMailbox => HasDetails && !string.Equals(RecipientTypeDetails, "SharedMailbox", StringComparison.OrdinalIgnoreCase);
    public bool CanConvertToRegularMailbox => HasDetails && string.Equals(RecipientTypeDetails, "SharedMailbox", StringComparison.OrdinalIgnoreCase);
    public bool HasConversionActions => CanConvertToSharedMailbox || CanConvertToRegularMailbox;
    public bool HasPendingMailboxChanges => SettingsEditor.HasPendingChanges || AutoReplyEditor.HasPendingChanges;
    public bool IsRetentionPolicyLoading => SettingsEditor.IsRetentionPolicyLoading;

    public ICommand RefreshCommand { get; }
    public ICommand RefreshRetentionPoliciesCommand { get; }
    public ICommand BackCommand { get; }
    public ICommand SavePermissionsCommand { get; }
    public ICommand DiscardPermissionsCommand { get; }
    public ICommand SaveMailboxChangesCommand { get; }
    public ICommand DiscardMailboxChangesCommand { get; }
    public ICommand ConvertToSharedMailboxCommand { get; }
    public ICommand ConvertToRegularMailboxCommand { get; }
    public ICommand AddPermissionCommand { get; }
    public ICommand RemovePermissionCommand { get; }
    public ICommand ModifyAutoMappingCommand { get; }
    public ICommand RefreshFolderPermissionsCommand => _permissionsEditor.RefreshFolderPermissionsCommand;
    public ICommand AddFolderPermissionCommand => _permissionsEditor.AddFolderPermissionCommand;
    public ICommand UpdateFolderPermissionCommand => _permissionsEditor.UpdateFolderPermissionCommand;
    public ICommand RemoveFolderPermissionCommand => _permissionsEditor.RemoveFolderPermissionCommand;

    private void OnPermissionsEditorPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (!string.IsNullOrWhiteSpace(e.PropertyName))
        {
            OnPropertyChanged(e.PropertyName);
        }
    }

    private void OnMailboxEditorPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(MailboxSettingsEditorViewModel.IsRetentionPolicyLoading))
        {
            OnPropertyChanged(nameof(IsRetentionPolicyLoading));
        }

        if (e.PropertyName is nameof(MailboxSettingsEditorViewModel.HasPendingChanges) or nameof(MailboxAutoReplyEditorViewModel.HasPendingChanges))
        {
            OnPropertyChanged(nameof(HasPendingMailboxChanges));
            CommandManager.InvalidateRequerySuggested();
        }
    }

    private void OnSelectedIdentityChanged(object? sender, string? identity)
    {
        if (_navigationService.CurrentPage == NavigationPage.Mailboxes)
        {
            Identity = identity;
            Details = null;
            ErrorMessage = null;
            _permissionsEditor.ClearPendingChanges();
            _permissionsEditor.ResetFolderPermissions();
            SettingsEditor.Reset();
            AutoReplyEditor.Reset();
            Licenses.Reset();
            RestoreMailbox.Reset(identity);
            OnPropertyChanged(nameof(HasPendingMailboxChanges));

            if (!string.IsNullOrEmpty(identity))
            {
                SafeLoadAsync(identity);
            }
        }
    }

    private async void SafeLoadAsync(string identity)
    {
        try
        {
            await LoadAsync(identity);
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Load failed: {ex.Message}");
        }
    }

    public async Task LoadAsync(string identity)
    {
        Identity = identity;
        await RefreshAsync(CancellationToken.None);
    }

    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrEmpty(Identity))
        {
            return;
        }

        var identitySnapshot = Identity;
        _loadCts?.Cancel();
        _loadCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

        IsLoading = true;
        ErrorMessage = null;
        _permissionsEditor.ClearPendingChanges();

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

            if (!string.Equals(Identity, identitySnapshot, StringComparison.Ordinal))
            {
                return;
            }

            if (result.IsSuccess && result.Value != null)
            {
                Details = result.Value;
                await _permissionsEditor.LoadFolderPermissionsAsync(_loadCts.Token);
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? "Failed to load mailbox details";
                _shellViewModel.AddLog(LogLevel.Error, $"Mailbox details load failed: {ErrorMessage}");
            }

            await SettingsEditor.LoadRetentionPoliciesAsync(_loadCts.Token);
            await Licenses.LoadAvailableLicensesAsync(_shellViewModel.IsExchangeConnected, _loadCts.Token);
        }
        catch (OperationCanceledException)
        {
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

    private void DiscardMailboxChanges()
    {
        SettingsEditor.DiscardChanges();
        AutoReplyEditor.DiscardChanges();
        _shellViewModel.AddLog(LogLevel.Information, "Mailbox changes discarded");
    }

    private async Task SaveMailboxChangesAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrEmpty(Identity))
        {
            return;
        }

        IsSaving = true;
        ErrorMessage = null;

        try
        {
            var settingsRequest = SettingsEditor.BuildUpdateRequest(Identity, out var settingsChanged, out var retentionPolicyChanged, out var retentionPolicyOverride);

            if (settingsChanged)
            {
                _shellViewModel.AddLog(LogLevel.Information, "Saving mailbox settings...");

                var result = await _workerService.UpdateMailboxSettingsAsync(settingsRequest, cancellationToken: cancellationToken);
                if (!result.IsSuccess)
                {
                    ErrorMessage = result.Error?.Message ?? "Unable to update mailbox settings.";
                    _shellViewModel.AddLog(LogLevel.Error, $"Mailbox settings update failed: {ErrorMessage}");
                    return;
                }
            }

            if (retentionPolicyChanged)
            {
                _shellViewModel.AddLog(LogLevel.Information, "Saving retention policy...");

                var retentionResult = await _workerService.SetRetentionPolicyAsync(
                    new SetRetentionPolicyRequest
                    {
                        Identity = Identity,
                        PolicyName = string.IsNullOrWhiteSpace(retentionPolicyOverride) ? null : retentionPolicyOverride
                    },
                    cancellationToken: cancellationToken);

                if (!retentionResult.IsSuccess)
                {
                    ErrorMessage = retentionResult.Error?.Message ?? "Unable to update the retention policy.";
                    _shellViewModel.AddLog(LogLevel.Error, $"Retention policy update failed: {ErrorMessage}");
                    return;
                }
            }

            var autoReplyRequest = AutoReplyEditor.BuildRequest(Identity, out var autoReplyValidationError);
            if (!string.IsNullOrWhiteSpace(autoReplyValidationError))
            {
                ErrorMessage = autoReplyValidationError;
                return;
            }

            var autoReplyChanged = autoReplyRequest != null;
            if ((settingsChanged || retentionPolicyChanged || autoReplyChanged) &&
                !ConfirmMutation(
                    "Update mailbox configuration",
                    $"{DisplayName} ({Identity})",
                    "Update mailbox settings, retention policy, or automatic replies.",
                    "Confirm mailbox save"))
            {
                return;
            }

            if (autoReplyChanged)
            {
                _shellViewModel.AddLog(LogLevel.Information, "Saving automatic replies...");

                var autoReplyResult = await _workerService.SetMailboxAutoReplyConfigurationAsync(
                    autoReplyRequest!,
                    cancellationToken: cancellationToken);

                if (!autoReplyResult.IsSuccess)
                {
                    ErrorMessage = autoReplyResult.Error?.Message ?? "Unable to update automatic replies.";
                    _shellViewModel.AddLog(LogLevel.Error, $"Auto-reply update failed: {ErrorMessage}");
                    return;
                }
            }

            if (settingsChanged || retentionPolicyChanged || autoReplyChanged)
            {
                await RefreshAsync(cancellationToken);

                if (retentionPolicyChanged)
                {
                    SettingsEditor.ApplyRetentionPolicyOverride(retentionPolicyOverride ?? string.Empty);
                }
            }
        }
        catch (OperationCanceledException)
        {
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

        if (!ConfirmMutation(
                "Convert mailbox to shared",
                $"{DisplayName} ({Identity})",
                "The mailbox becomes shared and licensing or access behavior can change.",
                "Confirm mailbox conversion"))
        {
            return;
        }

        IsSaving = true;
        ErrorMessage = null;

        try
        {
            var result = await _workerService.ConvertMailboxToSharedAsync(
                new ConvertMailboxToSharedRequest { Identity = Identity },
                cancellationToken: cancellationToken);

            if (result.IsSuccess)
            {
                _shellViewModel.AddLog(LogLevel.Information, $"Mailbox {Identity} converted to shared");
                await RefreshAsync(cancellationToken);
            }
            else
            {
                ErrorMessage = result.Error?.Message ?? "Unable to convert the mailbox.";
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

    private async Task ConvertToRegularMailboxAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrEmpty(Identity) || !CanConvertToRegularMailbox)
        {
            return;
        }

        if (!ConfirmMutation(
                "Convert mailbox to user",
                $"{DisplayName} ({Identity})",
                "The mailbox becomes a user mailbox and may require different licensing or policies.",
                "Confirm mailbox conversion"))
        {
            return;
        }

        IsSaving = true;
        ErrorMessage = null;

        try
        {
            var result = await _workerService.ConvertMailboxToRegularAsync(
                new ConvertMailboxToRegularRequest { Identity = Identity },
                cancellationToken: cancellationToken);

            if (result.IsSuccess)
            {
                _shellViewModel.AddLog(LogLevel.Information, $"Mailbox {Identity} converted to user");
                await RefreshAsync(cancellationToken);
            }
            else
            {
                ErrorMessage = result.Error?.Message ?? "Unable to convert the mailbox.";
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

    private void GoBack()
    {
        _navigationService.ClearSelection();
    }

    private string BuildQuotaUsageSummary()
    {
        if (Statistics == null)
        {
            return "Not available";
        }

        if (Statistics.TotalItemSizeBytes == null)
        {
            return Statistics.TotalItemSize ?? "Not available";
        }

        var quotaBytes = GetEffectiveQuotaBytes();
        if (quotaBytes == null || quotaBytes.Value <= 0)
        {
            return Statistics.TotalItemSize ?? "Not available";
        }

        var percent = Statistics.TotalItemSizeBytes.Value / (double)quotaBytes.Value * 100;
        var quotaLabel = GetEffectiveQuotaLabel();
        var suffix = string.IsNullOrWhiteSpace(quotaLabel) ? string.Empty : $" of {quotaLabel}";
        return $"{Statistics.TotalItemSize} ({percent:0.0}%{suffix})";
    }

    private long? GetEffectiveQuotaBytes()
    {
        if (Features == null)
        {
            return null;
        }

        return Features.ProhibitSendReceiveQuotaBytes
            ?? Features.ProhibitSendQuotaBytes
            ?? Features.IssueWarningQuotaBytes;
    }

    private string? GetEffectiveQuotaLabel()
    {
        if (Features == null)
        {
            return null;
        }

        return Features.ProhibitSendReceiveQuota
            ?? Features.ProhibitSendQuota
            ?? Features.IssueWarningQuota;
    }
}

