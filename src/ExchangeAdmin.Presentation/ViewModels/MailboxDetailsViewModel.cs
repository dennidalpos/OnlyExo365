using System.Collections.ObjectModel;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

             
                                                        
              
public class MailboxDetailsViewModel : ViewModelBase
{
    private readonly IWorkerService _workerService;
    private readonly NavigationService _navigationService;
    private readonly ShellViewModel _shellViewModel;
    private readonly CacheService _cacheService;

    private CancellationTokenSource? _loadCts;
    private static readonly TimeSpan RetentionPolicyCacheTtl = TimeSpan.FromMinutes(10);

    private bool _isLoading;
    private bool _isSaving;
    private string? _errorMessage;
    private string? _identity;
    private MailboxDetailsDto? _details;
    private bool _isRetentionPolicyLoading;

                                 
    private readonly List<PermissionDeltaActionDto> _pendingActions = new();
    private bool _hasPendingChanges;

                               
    private bool _hasPendingMailboxChanges;
    private bool _isInitializingMailboxSettings;
    private bool _isRefreshingRetentionPolicies;
    private MailboxSettingsSnapshot? _originalMailboxSettings;
    private string? _retentionPolicyFallback;

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

    private readonly ObservableCollection<RetentionPolicySummaryDto> _availableRetentionPolicies = new();
    private string? _selectedRetentionPolicy;

    // License management
    private bool _isLicenseLoading;
    private bool _isLicenseSaving;
    private string? _licenseErrorMessage;
    private readonly ObservableCollection<UserLicenseDto> _assignedLicenses = new();
    private readonly ObservableCollection<TenantLicenseDto> _availableLicenses = new();
    private TenantLicenseDto? _selectedLicenseToAdd;

    private string? _restoreSourceIdentity;
    private string? _restoreTargetMailbox;
    private bool _restoreAllowLegacyDnMismatch;
    private bool _isRestoringMailbox;
    private RestoreMailboxResponse? _restoreMailboxResponse;
    private string? _restoreMailboxErrorMessage;

    private string? _newPermissionUser;
    private PermissionType _newPermissionType = PermissionType.FullAccess;
    private bool _newPermissionAutoMapping = true;

    public MailboxDetailsViewModel(IWorkerService workerService, NavigationService navigationService, ShellViewModel shellViewModel, CacheService cacheService)
    {
        _workerService = workerService;
        _navigationService = navigationService;
        _shellViewModel = shellViewModel;
        _cacheService = cacheService;

                                       
        _navigationService.SelectedIdentityChanged += OnSelectedIdentityChanged;

                   
        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
        RefreshRetentionPoliciesCommand = new AsyncRelayCommand(LoadRetentionPoliciesAsync, () => !IsRetentionPolicyLoading);
        BackCommand = new RelayCommand(GoBack);
        SavePermissionsCommand = new AsyncRelayCommand(SavePermissionsAsync, () => HasPendingChanges && !IsSaving);
        DiscardPermissionsCommand = new RelayCommand(DiscardPendingChanges, () => HasPendingChanges);

        SaveMailboxChangesCommand = new AsyncRelayCommand(SaveMailboxChangesAsync, () => HasPendingMailboxChanges && !IsSaving);
        DiscardMailboxChangesCommand = new RelayCommand(DiscardMailboxChanges, () => HasPendingMailboxChanges);
        ConvertToSharedMailboxCommand = new AsyncRelayCommand(ConvertToSharedMailboxAsync, () => CanConvertToSharedMailbox && !IsSaving);
        ConvertToRegularMailboxCommand = new AsyncRelayCommand(ConvertToRegularMailboxAsync, () => CanConvertToRegularMailbox && !IsSaving);

        AddPermissionCommand = new RelayCommand(AddPermission, () => CanAddPermission);
        RemovePermissionCommand = new RelayCommand<object>(RemovePermission);
        ModifyAutoMappingCommand = new RelayCommand<object>(ModifyAutoMapping);

        AddLicenseCommand = new AsyncRelayCommand(AddLicenseAsync, () => CanAddLicense && !IsLicenseSaving);
        RemoveLicenseCommand = new AsyncRelayCommand<UserLicenseDto>(RemoveLicenseAsync, _ => !IsLicenseSaving);
        RefreshLicensesCommand = new AsyncRelayCommand(RefreshLicensesAsync, () => !IsLicenseLoading);

        RestoreMailboxCommand = new AsyncRelayCommand(RestoreMailboxAsync, () => !IsRestoringMailbox);
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

    public string? RestoreSourceIdentity
    {
        get => _restoreSourceIdentity;
        set => SetProperty(ref _restoreSourceIdentity, value);
    }

    public string? RestoreTargetMailbox
    {
        get => _restoreTargetMailbox;
        set => SetProperty(ref _restoreTargetMailbox, value);
    }

    public bool RestoreAllowLegacyDnMismatch
    {
        get => _restoreAllowLegacyDnMismatch;
        set => SetProperty(ref _restoreAllowLegacyDnMismatch, value);
    }

    public bool IsRestoringMailbox
    {
        get => _isRestoringMailbox;
        private set
        {
            if (SetProperty(ref _isRestoringMailbox, value))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public RestoreMailboxResponse? RestoreMailboxResponse
    {
        get => _restoreMailboxResponse;
        private set
        {
            if (SetProperty(ref _restoreMailboxResponse, value))
            {
                _restoreMailboxErrorMessage = null;
                OnPropertyChanged(nameof(HasRestoreMailboxResponse));
                OnPropertyChanged(nameof(RestoreScenarioText));
                OnPropertyChanged(nameof(RestoreActionText));
                OnPropertyChanged(nameof(RestoreStatusText));
                OnPropertyChanged(nameof(RestoreStatusDetail));
                OnPropertyChanged(nameof(RestorePercentComplete));
                OnPropertyChanged(nameof(RestoreRequestGuid));
                OnPropertyChanged(nameof(RestoreErrorCodeText));
                OnPropertyChanged(nameof(RestoreErrorMessage));
                OnPropertyChanged(nameof(HasRestoreError));
                OnPropertyChanged(nameof(RestoreProgressValue));
                OnPropertyChanged(nameof(IsRestoreProgressIndeterminate));
                OnPropertyChanged(nameof(RestoreProgressText));
            }
        }
    }

    public bool HasRestoreMailboxResponse => RestoreMailboxResponse != null;

    public string RestoreScenarioText => RestoreMailboxResponse == null
        ? "-"
        : RestoreMailboxResponse.Scenario switch
        {
            RestoreMailboxScenario.SoftDeleted => "Soft-deleted",
            RestoreMailboxScenario.HardDeleted => "Hard-deleted (inactive)",
            RestoreMailboxScenario.Existing => "Mailbox esistente",
            RestoreMailboxScenario.NotFound => "Non trovata",
            _ => "Sconosciuto"
        };

    public string RestoreActionText => RestoreMailboxResponse?.Action ?? "-";

    public string RestoreStatusText => RestoreMailboxResponse == null
        ? "-"
        : RestoreMailboxResponse.Status switch
        {
            RestoreMailboxStatus.InProgress => "In corso",
            RestoreMailboxStatus.Completed => "Completato",
            RestoreMailboxStatus.Failed => "Errore",
            _ => "Non avviato"
        };

    public string RestoreStatusDetail => RestoreMailboxResponse?.StatusDetail ?? "-";

    public int? RestorePercentComplete => RestoreMailboxResponse?.PercentComplete;

    public string RestoreRequestGuid => RestoreMailboxResponse?.RequestGuid ?? "-";

    public string RestoreErrorCodeText => RestoreMailboxResponse?.Error?.Code.ToString() ?? "-";

    public string? RestoreErrorMessage
    {
        get => RestoreMailboxResponse?.Error?.Message ?? _restoreMailboxErrorMessage;
        private set
        {
            if (SetProperty(ref _restoreMailboxErrorMessage, value))
            {
                OnPropertyChanged(nameof(HasRestoreError));
            }
        }
    }

    public bool HasRestoreError => !string.IsNullOrWhiteSpace(RestoreErrorMessage);

    public int RestoreProgressValue => RestoreMailboxResponse?.PercentComplete ?? 0;

    public bool IsRestoreProgressIndeterminate =>
        RestoreMailboxResponse?.Status == RestoreMailboxStatus.InProgress &&
        (RestoreMailboxResponse?.PercentComplete.HasValue == false);

    public string RestoreProgressText => RestoreMailboxResponse?.PercentComplete.HasValue == true
        ? $"{RestoreMailboxResponse.PercentComplete}% completato"
        : "Avanzamento non disponibile";

    public string? Identity
    {
        get => _identity;
        private set => SetProperty(ref _identity, value);
    }

    // Batched property names for Details change notification
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
                if (!string.IsNullOrWhiteSpace(_details?.RetentionPolicy))
                {
                    _retentionPolicyFallback = _details.RetentionPolicy;
                }

                // Batch notify all related properties
                foreach (var prop in DetailsRelatedProperties)
                {
                    OnPropertyChanged(prop);
                }

                UpdatePermissionsDisplay();
                UpdateAssignedLicensesDisplay();
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
    public string QuotaUsageSummary => BuildQuotaUsageSummary();

    public ObservableCollection<RetentionPolicySummaryDto> AvailableRetentionPolicies => _availableRetentionPolicies;

    public string? SelectedRetentionPolicy
    {
        get => _selectedRetentionPolicy;
        set
        {
            if (SetProperty(ref _selectedRetentionPolicy, value))
            {
                UpdatePendingMailboxChanges();
                OnPropertyChanged(nameof(SelectedRetentionPolicyDescription));
                OnPropertyChanged(nameof(SelectedRetentionPolicyRequiresArchive));
                OnPropertyChanged(nameof(ShowArchiveRequiredWarning));
            }
        }
    }

    public string? SelectedRetentionPolicyDescription
        => AvailableRetentionPolicies.FirstOrDefault(policy => IsPolicySelected(policy))?.Description;

    public bool SelectedRetentionPolicyRequiresArchive
        => AvailableRetentionPolicies.FirstOrDefault(policy => IsPolicySelected(policy))?.RequiresArchive ?? false;

    public bool ShowArchiveRequiredWarning => !ArchiveEnabled && SelectedRetentionPolicyRequiresArchive;

    public bool IsRetentionPolicyLoading
    {
        get => _isRetentionPolicyLoading;
        private set
        {
            if (SetProperty(ref _isRetentionPolicyLoading, value))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    // License properties
    public ObservableCollection<UserLicenseDto> AssignedLicenses => _assignedLicenses;
    public ObservableCollection<TenantLicenseDto> AvailableLicenses => _availableLicenses;
    public bool HasAssignedLicenses => _assignedLicenses.Count > 0;

    public TenantLicenseDto? SelectedLicenseToAdd
    {
        get => _selectedLicenseToAdd;
        set
        {
            if (SetProperty(ref _selectedLicenseToAdd, value))
            {
                OnPropertyChanged(nameof(CanAddLicense));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool IsLicenseLoading
    {
        get => _isLicenseLoading;
        private set
        {
            if (SetProperty(ref _isLicenseLoading, value))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool IsLicenseSaving
    {
        get => _isLicenseSaving;
        private set
        {
            if (SetProperty(ref _isLicenseSaving, value))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string? LicenseErrorMessage
    {
        get => _licenseErrorMessage;
        private set
        {
            if (SetProperty(ref _licenseErrorMessage, value))
            {
                OnPropertyChanged(nameof(HasLicenseError));
            }
        }
    }

    public bool HasLicenseError => !string.IsNullOrEmpty(LicenseErrorMessage);
    public bool CanAddLicense => SelectedLicenseToAdd != null && !string.IsNullOrEmpty(PrimarySmtpAddress);

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
                OnPropertyChanged(nameof(ShowArchiveRequiredWarning));
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

                          
    public ObservableCollection<PermissionDisplayItem> FullAccessPermissions { get; } = new();
    public ObservableCollection<PermissionDisplayItem> SendAsPermissions { get; } = new();
    public ObservableCollection<PermissionDisplayItem> SendOnBehalfPermissions { get; } = new();

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

    public bool CanRefresh => !IsLoading && _shellViewModel.IsExchangeConnected && !string.IsNullOrEmpty(Identity);
    public bool CanAddPermission => !string.IsNullOrWhiteSpace(NewPermissionUser);
    public bool CanConvertToSharedMailbox => HasDetails && !string.Equals(RecipientTypeDetails, "SharedMailbox", StringComparison.OrdinalIgnoreCase);
    public bool CanConvertToRegularMailbox => HasDetails && string.Equals(RecipientTypeDetails, "SharedMailbox", StringComparison.OrdinalIgnoreCase);
    public bool HasConversionActions => CanConvertToSharedMailbox || CanConvertToRegularMailbox;

    #endregion

    #region Commands

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
    public ICommand AddLicenseCommand { get; }
    public ICommand RemoveLicenseCommand { get; }
    public ICommand RefreshLicensesCommand { get; }
    public ICommand RestoreMailboxCommand { get; }

    #endregion

    #region Methods

    private void OnSelectedIdentityChanged(object? sender, string? identity)
    {
        if (_navigationService.CurrentPage == NavigationPage.Mailboxes ||
            _navigationService.CurrentPage == NavigationPage.SharedMailboxes)
        {
            Identity = identity;
            Details = null;
            _retentionPolicyFallback = null;
            _originalMailboxSettings = null;
            HasPendingMailboxChanges = false;
            ClearPendingActions();
            ResetRestoreMailboxState(identity);
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
        if (string.IsNullOrEmpty(Identity)) return;

        var identitySnapshot = Identity;
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

            if (!string.Equals(Identity, identitySnapshot, StringComparison.Ordinal))
            {
                return;
            }

            if (result.IsSuccess && result.Value != null)
            {
                Details = result.Value;
            }
            else if (result.WasCancelled)
            {
                         
            }
            else
            {
                ErrorMessage = result.Error?.Message ?? "Failed to load mailbox details";
                _shellViewModel.AddLog(LogLevel.Error, $"Mailbox details load failed: {ErrorMessage}");
            }

            await LoadRetentionPoliciesAsync(_loadCts.Token);
            await LoadAvailableLicensesAsync(_loadCts.Token);
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

    private async Task LoadRetentionPoliciesAsync(CancellationToken cancellationToken)
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            return;
        }

        IsRetentionPolicyLoading = true;
        var selectedPolicySnapshot = SelectedRetentionPolicy;

        try
        {
            _isRefreshingRetentionPolicies = true;

            // Try to load from cache first
            var cachedPolicies = _cacheService.Get<List<RetentionPolicySummaryDto>>(CacheService.Keys.RetentionPolicies);

            List<RetentionPolicySummaryDto>? policies = null;

            if (cachedPolicies != null)
            {
                policies = cachedPolicies;
            }
            else
            {
                var result = await _workerService.GetRetentionPoliciesAsync(
                    new GetRetentionPoliciesRequest(),
                    cancellationToken: cancellationToken);

                if (result.IsSuccess && result.Value != null)
                {
                    policies = result.Value.Policies.OrderBy(policy => policy.Name).ToList();
                    // Cache the policies
                    _cacheService.Set(CacheService.Keys.RetentionPolicies, policies, RetentionPolicyCacheTtl);
                }
                else if (!result.WasCancelled)
                {
                    _shellViewModel.AddLog(LogLevel.Warning, result.Error?.Message ?? "Impossibile recuperare le retention policy");
                }
            }

            if (policies != null)
            {
                RunOnUiThread(() =>
                {
                    AvailableRetentionPolicies.Clear();
                    AvailableRetentionPolicies.Add(new RetentionPolicySummaryDto
                    {
                        Name = string.Empty,
                        Description = "Nessuna policy assegnata"
                    });

                    foreach (var policy in policies)
                    {
                        AvailableRetentionPolicies.Add(policy);
                    }

                    if (!string.IsNullOrWhiteSpace(selectedPolicySnapshot))
                    {
                        SelectedRetentionPolicy = selectedPolicySnapshot;
                    }

                    OnPropertyChanged(nameof(SelectedRetentionPolicyDescription));
                    OnPropertyChanged(nameof(SelectedRetentionPolicyRequiresArchive));
                    OnPropertyChanged(nameof(ShowArchiveRequiredWarning));
                });
            }
        }
        catch (OperationCanceledException)
        {
            // Cancelled
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Retention policies load failed: {ex.Message}");
        }
        finally
        {
            _isRefreshingRetentionPolicies = false;
            UpdatePendingMailboxChanges();
            IsRetentionPolicyLoading = false;
        }
    }

    private void UpdateAssignedLicensesDisplay()
    {
        _assignedLicenses.Clear();
        if (Details?.AssignedLicenses != null)
        {
            foreach (var lic in Details.AssignedLicenses)
            {
                _assignedLicenses.Add(lic);
            }
        }
        OnPropertyChanged(nameof(HasAssignedLicenses));
    }

    private async Task LoadAvailableLicensesAsync(CancellationToken cancellationToken)
    {
        if (!_shellViewModel.IsExchangeConnected) return;

        IsLicenseLoading = true;
        try
        {
            var result = await _workerService.GetAvailableLicensesAsync(cancellationToken: cancellationToken);
            if (result.IsSuccess && result.Value != null)
            {
                RunOnUiThread(() =>
                {
                    _availableLicenses.Clear();
                    foreach (var lic in result.Value.Licenses.OrderBy(l => l.SkuPartNumber))
                    {
                        _availableLicenses.Add(lic);
                    }
                });
            }
            else if (!result.WasCancelled)
            {
                _shellViewModel.AddLog(LogLevel.Warning, result.Error?.Message ?? "Impossibile recuperare le licenze disponibili");
            }
        }
        catch (OperationCanceledException) { }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Available licenses load failed: {ex.Message}");
        }
        finally
        {
            IsLicenseLoading = false;
        }
    }

    private async Task RefreshLicensesAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrEmpty(PrimarySmtpAddress)) return;

        IsLicenseLoading = true;
        LicenseErrorMessage = null;
        try
        {
            var request = new GetUserLicensesRequest { UserPrincipalName = PrimarySmtpAddress };
            var result = await _workerService.GetUserLicensesAsync(request, cancellationToken: cancellationToken);
            if (result.IsSuccess && result.Value != null)
            {
                RunOnUiThread(() =>
                {
                    _assignedLicenses.Clear();
                    foreach (var lic in result.Value.Licenses)
                    {
                        _assignedLicenses.Add(lic);
                    }
                    OnPropertyChanged(nameof(HasAssignedLicenses));
                });
            }
            else if (!result.WasCancelled)
            {
                LicenseErrorMessage = result.Error?.Message ?? "Impossibile recuperare le licenze utente";
            }

            await LoadAvailableLicensesAsync(cancellationToken);
        }
        catch (OperationCanceledException) { }
        catch (Exception ex)
        {
            LicenseErrorMessage = ex.Message;
        }
        finally
        {
            IsLicenseLoading = false;
        }
    }

    private async Task AddLicenseAsync(CancellationToken cancellationToken)
    {
        if (SelectedLicenseToAdd == null || string.IsNullOrEmpty(PrimarySmtpAddress)) return;

        IsLicenseSaving = true;
        LicenseErrorMessage = null;
        try
        {
            var request = new SetUserLicenseRequest
            {
                UserPrincipalName = PrimarySmtpAddress,
                AddLicenseSkuIds = new List<string> { SelectedLicenseToAdd.SkuId }
            };

            _shellViewModel.AddLog(LogLevel.Information, $"Assegnazione licenza {SelectedLicenseToAdd.SkuPartNumber}...");

            var result = await _workerService.SetUserLicenseAsync(request, cancellationToken: cancellationToken);
            if (result.IsSuccess)
            {
                _shellViewModel.AddLog(LogLevel.Information, $"Licenza {SelectedLicenseToAdd.SkuPartNumber} assegnata");
                SelectedLicenseToAdd = null;
                await RefreshLicensesAsync(cancellationToken);
            }
            else if (!result.WasCancelled)
            {
                LicenseErrorMessage = result.Error?.Message ?? "Impossibile assegnare la licenza";
            }
        }
        catch (OperationCanceledException) { }
        catch (Exception ex)
        {
            LicenseErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Add license error: {ex.Message}");
        }
        finally
        {
            IsLicenseSaving = false;
        }
    }

    private async Task RemoveLicenseAsync(UserLicenseDto? license, CancellationToken cancellationToken)
    {
        if (license == null || string.IsNullOrEmpty(PrimarySmtpAddress)) return;

        var confirm = System.Windows.MessageBox.Show(
            $"Rimuovere la licenza {license.SkuPartNumber} da {DisplayName}?",
            "Conferma rimozione licenza",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Warning);

        if (confirm != System.Windows.MessageBoxResult.Yes) return;

        IsLicenseSaving = true;
        LicenseErrorMessage = null;
        try
        {
            var request = new SetUserLicenseRequest
            {
                UserPrincipalName = PrimarySmtpAddress,
                RemoveLicenseSkuIds = new List<string> { license.SkuId }
            };

            _shellViewModel.AddLog(LogLevel.Information, $"Rimozione licenza {license.SkuPartNumber}...");

            var result = await _workerService.SetUserLicenseAsync(request, cancellationToken: cancellationToken);
            if (result.IsSuccess)
            {
                _shellViewModel.AddLog(LogLevel.Information, $"Licenza {license.SkuPartNumber} rimossa");
                await RefreshLicensesAsync(cancellationToken);
            }
            else if (!result.WasCancelled)
            {
                LicenseErrorMessage = result.Error?.Message ?? "Impossibile rimuovere la licenza";
            }
        }
        catch (OperationCanceledException) { }
        catch (Exception ex)
        {
            LicenseErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Remove license error: {ex.Message}");
        }
        finally
        {
            IsLicenseSaving = false;
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
            var displayName = string.IsNullOrWhiteSpace(user.DisplayName) ? user.Identity : user.DisplayName;
            SendOnBehalfPermissions.Add(new PermissionDisplayItem
            {
                User = displayName,
                Identity = user.Identity,
                PermissionType = PermissionType.SendOnBehalf
            });
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
                SendOnBehalfPermissions.Add(new PermissionDisplayItem
                {
                    User = NewPermissionUser.Trim(),
                    Identity = NewPermissionUser.Trim(),
                    PermissionType = PermissionType.SendOnBehalf,
                    IsPending = true
                });
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
            user = string.IsNullOrWhiteSpace(displayItem.Identity) ? displayItem.User : displayItem.Identity;

                                  
            switch (permType)
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
            User = string.IsNullOrWhiteSpace(displayItem.Identity) ? displayItem.User : displayItem.Identity,
            AutoMapping = newAutoMapping,
            Description = $"Set AutoMapping to {newAutoMapping} for {displayItem.User}"
        };

        _pendingActions.Add(action);
        HasPendingChanges = _pendingActions.Count > 0;

                          
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

                                                     
                await RefreshAsync(cancellationToken);
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? "Failed to apply permissions";
            }
        }
        catch (OperationCanceledException)
        {
                     
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
        UpdatePermissionsDisplay();                          
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
        var retentionPolicy = !string.IsNullOrWhiteSpace(Details.RetentionPolicy)
            ? Details.RetentionPolicy
            : _retentionPolicyFallback ?? string.Empty;
        SelectedRetentionPolicy = retentionPolicy;

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
            AutoReplyInternalMessage = NormalizeAutoReplyMessage(autoReply.InternalMessage);
            AutoReplyExternalMessage = NormalizeAutoReplyMessage(autoReply.ExternalMessage);
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
        if (_isInitializingMailboxSettings || _isRefreshingRetentionPolicies || _originalMailboxSettings == null)
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
        SelectedRetentionPolicy = _originalMailboxSettings.RetentionPolicy;
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
            var normalizedRetentionPolicy = NormalizeInput(SelectedRetentionPolicy);

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

            var retentionPolicyChanged = !string.Equals(normalizedRetentionPolicy, _originalMailboxSettings.RetentionPolicy, StringComparison.Ordinal);

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

            string? retentionPolicyOverride = null;

            if (retentionPolicyChanged)
            {
                _shellViewModel.AddLog(LogLevel.Information, "Salvataggio retention policy...");

                var retentionPolicyRequest = new SetRetentionPolicyRequest
                {
                    Identity = Identity,
                    PolicyName = string.IsNullOrWhiteSpace(normalizedRetentionPolicy) ? null : normalizedRetentionPolicy
                };

                var retentionPolicyResult = await _workerService.SetRetentionPolicyAsync(retentionPolicyRequest, cancellationToken: cancellationToken);

                if (!retentionPolicyResult.IsSuccess)
                {
                    ErrorMessage = retentionPolicyResult.Error?.Message ?? "Impossibile aggiornare la retention policy";
                    _shellViewModel.AddLog(LogLevel.Error, $"Retention policy update failed: {ErrorMessage}");
                    return;
                }

                retentionPolicyOverride = normalizedRetentionPolicy;
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

                var normalizedInternalMessage = NormalizeAutoReplyMessage(AutoReplyInternalMessage);
                var normalizedExternalMessage = NormalizeAutoReplyMessage(AutoReplyExternalMessage);

                var autoReplyRequest = new SetMailboxAutoReplyConfigurationRequest
                {
                    Identity = Identity,
                    AutoReplyState = AutoReplyEnabled
                        ? AutoReplyScheduled ? "Scheduled" : "Enabled"
                        : "Disabled",
                    StartTime = AutoReplyScheduled ? BuildDateTime(AutoReplyStartDate, AutoReplyStartTime) : null,
                    EndTime = AutoReplyScheduled ? BuildDateTime(AutoReplyEndDate, AutoReplyEndTime) : null,
                    InternalMessage = normalizedInternalMessage,
                    ExternalMessage = normalizedExternalMessage,
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

            if (settingsChanged || autoReplyChanged || retentionPolicyChanged)
            {
                HasPendingMailboxChanges = false;
                CommandManager.InvalidateRequerySuggested();
                await RefreshAsync(cancellationToken);

                if (retentionPolicyOverride != null)
                {
                    ApplyRetentionPolicyOverride(retentionPolicyOverride);
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

    private async Task ConvertToRegularMailboxAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrEmpty(Identity) || !CanConvertToRegularMailbox)
        {
            return;
        }

        var confirm = System.Windows.MessageBox.Show(
            $"Convertire {DisplayName} in mailbox utente?",
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
            var request = new ConvertMailboxToRegularRequest
            {
                Identity = Identity
            };

            var result = await _workerService.ConvertMailboxToRegularAsync(request, cancellationToken: cancellationToken);

            if (result.IsSuccess)
            {
                _shellViewModel.AddLog(LogLevel.Information, $"Mailbox {Identity} convertita in regolare");
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
            RetentionPolicy = NormalizeInput(SelectedRetentionPolicy),
            AutoReplyEnabled = AutoReplyEnabled,
            AutoReplyScheduled = AutoReplyScheduled,
            AutoReplyStartDate = BuildDateTime(AutoReplyStartDate, AutoReplyStartTime),
            AutoReplyEndDate = BuildDateTime(AutoReplyEndDate, AutoReplyEndTime),
            AutoReplyInternalMessage = NormalizeAutoReplyMessage(AutoReplyInternalMessage),
            AutoReplyExternalMessage = NormalizeAutoReplyMessage(AutoReplyExternalMessage),
            AutoReplyExternalAudience = NormalizeInput(AutoReplyExternalAudience)
        };
    }

    private void ApplyRetentionPolicyOverride(string policyName)
    {
        _isInitializingMailboxSettings = true;
        SelectedRetentionPolicy = policyName;
        _isInitializingMailboxSettings = false;
        _retentionPolicyFallback = policyName;

        if (Details != null)
        {
            Details.RetentionPolicy = string.IsNullOrWhiteSpace(policyName) ? null : policyName;
        }

        _originalMailboxSettings = CaptureMailboxSettingsSnapshot();
        HasPendingMailboxChanges = false;
        CommandManager.InvalidateRequerySuggested();
    }

    private void ResetRestoreMailboxState(string? identity)
    {
        RestoreSourceIdentity = identity;
        RestoreTargetMailbox = null;
        RestoreAllowLegacyDnMismatch = false;
        RestoreMailboxResponse = null;
        RestoreErrorMessage = null;
    }

    private async Task RestoreMailboxAsync(CancellationToken cancellationToken)
    {
        var sourceIdentity = RestoreSourceIdentity?.Trim();
        var targetMailbox = RestoreTargetMailbox?.Trim();

        if (string.IsNullOrWhiteSpace(sourceIdentity))
        {
            RestoreMailboxResponse = null;
            RestoreErrorMessage = "Specificare l'UPN o il GUID della mailbox sorgente.";
            return;
        }

        IsRestoringMailbox = true;
        RestoreMailboxResponse = null;
        RestoreErrorMessage = null;

        try
        {
            var request = new RestoreMailboxRequest
            {
                SourceIdentity = sourceIdentity,
                TargetMailbox = string.IsNullOrWhiteSpace(targetMailbox) ? null : targetMailbox,
                AllowLegacyDnMismatch = RestoreAllowLegacyDnMismatch
            };

            _shellViewModel.AddLog(LogLevel.Information, $"Avvio ripristino mailbox per {request.SourceIdentity}...");

            var result = await _workerService.RestoreMailboxAsync(request, cancellationToken: cancellationToken);

            if (result.IsSuccess && result.Value != null)
            {
                RestoreMailboxResponse = result.Value;
                if (RestoreMailboxResponse.Error != null)
                {
                    _shellViewModel.AddLog(LogLevel.Warning, $"Ripristino mailbox con errore: {RestoreMailboxResponse.Error.Message}");
                }
                else
                {
                    _shellViewModel.AddLog(LogLevel.Information, $"Ripristino mailbox avviato: {RestoreMailboxResponse.Status}");
                }
            }
            else
            {
                RestoreErrorMessage = result.Error?.Message ?? "Impossibile avviare il ripristino della mailbox.";
                if (result.IsSuccess && result.Value == null)
                {
                    RestoreErrorMessage = "Risposta ripristino non disponibile.";
                }
                _shellViewModel.AddLog(LogLevel.Error, $"Ripristino mailbox fallito: {RestoreErrorMessage}");
            }
        }
        catch (OperationCanceledException)
        {
                     
        }
        catch (Exception ex)
        {
            RestoreErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Ripristino mailbox error: {ex.Message}");
        }
        finally
        {
            IsRestoringMailbox = false;
        }
    }

    private string BuildQuotaUsageSummary()
    {
        if (Statistics == null)
        {
            return "Non disponibile";
        }

        if (Statistics.TotalItemSizeBytes == null)
        {
            return Statistics.TotalItemSize ?? "Non disponibile";
        }

        var quotaBytes = GetEffectiveQuotaBytes();
        if (quotaBytes == null || quotaBytes.Value <= 0)
        {
            return Statistics.TotalItemSize ?? "Non disponibile";
        }

        var percent = Statistics.TotalItemSizeBytes.Value / (double)quotaBytes.Value * 100;
        var quotaLabel = GetEffectiveQuotaLabel();
        var labelSuffix = string.IsNullOrWhiteSpace(quotaLabel) ? string.Empty : $" di {quotaLabel}";
        return $"{Statistics.TotalItemSize} ({percent:0.0}%{labelSuffix})";
    }

    private bool IsPolicySelected(RetentionPolicySummaryDto policy)
    {
        var selected = SelectedRetentionPolicy ?? string.Empty;
        return string.Equals(policy.Name, selected, StringComparison.OrdinalIgnoreCase);
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

    private static string NormalizeInput(string? value) => string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();

    private static string NormalizeAutoReplyMessage(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var stripped = StripHtml(value);
        return string.IsNullOrWhiteSpace(stripped) ? string.Empty : stripped.Trim();
    }

    private static string StripHtml(string value)
    {
        var normalized = Regex.Replace(value, @"<\s*br\s*/?>", "\n", RegexOptions.IgnoreCase);
        normalized = Regex.Replace(normalized, @"</\s*p\s*>", "\n", RegexOptions.IgnoreCase);
        normalized = Regex.Replace(normalized, @"</\s*div\s*>", "\n", RegexOptions.IgnoreCase);
        normalized = Regex.Replace(normalized, @"<\s*li\s*>", "- ", RegexOptions.IgnoreCase);
        normalized = Regex.Replace(normalized, @"</\s*li\s*>", "\n", RegexOptions.IgnoreCase);
        normalized = Regex.Replace(normalized, @"<[^>]+>", string.Empty);
        normalized = WebUtility.HtmlDecode(normalized);
        normalized = normalized.Replace("\r", string.Empty);
        normalized = Regex.Replace(normalized, @"\n{3,}", "\n\n");
        return normalized.Trim();
    }

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
    public string RetentionPolicy { get; set; } = string.Empty;
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
            && string.Equals(RetentionPolicy, other.RetentionPolicy, StringComparison.Ordinal)
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
        hash.Add(RetentionPolicy);
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

             
                            
              
public class PermissionDisplayItem
{
    public string User { get; set; } = string.Empty;
    public string Identity { get; set; } = string.Empty;
    public PermissionType PermissionType { get; set; }
    public bool? AutoMapping { get; set; }
    public bool IsInherited { get; set; }
    public bool IsPending { get; set; }
}
