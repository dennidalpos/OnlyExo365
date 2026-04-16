using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;
using ExchangeAdmin.Presentation.Helpers;

namespace ExchangeAdmin.Presentation.ViewModels;

public sealed class MailSecurityViewModel : ViewModelBase
{
    private readonly IMailSecurityWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;

    private bool _isLoadingWorkspace;
    private bool _isSaving;
    private string? _errorMessage;
    private string? _warningsText;

    private DkimSigningConfigDto? _selectedDkimConfig;
    private HostedContentFilterPolicyDto? _selectedAntiSpamPolicy;
    private AntiPhishPolicyDto? _selectedAntiPhishPolicy;
    private MalwareFilterPolicyDto? _selectedMalwarePolicy;
    private QuarantinePolicyDto? _selectedQuarantinePolicy;
    private HostedOutboundSpamFilterPolicyDto? _selectedOutboundSpamPolicy;

    private bool _dkimEnabled;

    private bool _antiSpamRuleEnabled;
    private int _antiSpamBulkThreshold = 6;
    private string _antiSpamSpamAction = "MoveToJmf";
    private string _antiSpamHighConfidenceSpamAction = "Quarantine";
    private string _antiSpamPhishSpamAction = "Quarantine";

    private bool _antiPhishRuleEnabled;
    private bool _antiPhishEnableSpoofIntelligence;
    private bool _antiPhishEnableMailboxIntelligence;
    private bool _antiPhishEnableTargetedUserProtection;
    private bool _antiPhishHonorDmarcPolicy = true;
    private int _antiPhishThresholdLevel = 2;
    private string _antiPhishMailboxIntelligenceProtectionAction = "Quarantine";
    private string _antiPhishTargetedUserProtectionAction = "Quarantine";
    private string _antiPhishAuthenticationFailAction = "MoveToJmf";
    private string _antiPhishDmarcRejectAction = "Quarantine";
    private string _antiPhishDmarcQuarantineAction = "Quarantine";

    private bool _malwareRuleEnabled;
    private bool _malwareEnableFileFilter;
    private bool _malwareZapEnabled = true;
    private string _malwareFileTypeAction = "Reject";

    private string _quarantinePermissions = "FullAccess";
    private bool _quarantineEsnEnabled = true;
    private int _quarantineEndUserSpamNotificationFrequency = 1;
    private int _quarantineRetentionDays = 15;
    private bool _quarantineOrganizationBrandingEnabled;

    private int _outboundRecipientLimitExternalPerHour = 500;
    private int _outboundRecipientLimitInternalPerHour = 1000;
    private int _outboundRecipientLimitPerDay = 1000;
    private string _outboundActionWhenThresholdReached = "BlockUser";
    private string _outboundAutoForwardingMode = "Automatic";

    public MailSecurityViewModel(IMailSecurityWorkerService workerService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;

        RefreshWorkspaceCommand = new AsyncRelayCommand(RefreshWorkspaceAsync, () => CanRefreshWorkspace);
        SaveDkimCommand = new AsyncRelayCommand(SaveDkimAsync, () => CanSaveDkim);
        SaveAntiSpamCommand = new AsyncRelayCommand(SaveAntiSpamAsync, () => CanSaveAntiSpam);
        SaveAntiPhishCommand = new AsyncRelayCommand(SaveAntiPhishAsync, () => CanSaveAntiPhish);
        SaveMalwareCommand = new AsyncRelayCommand(SaveMalwareAsync, () => CanSaveMalware);
        SaveQuarantineCommand = new AsyncRelayCommand(SaveQuarantineAsync, () => CanSaveQuarantine);
        SaveOutboundSpamCommand = new AsyncRelayCommand(SaveOutboundSpamAsync, () => CanSaveOutboundSpam);

        _shellViewModel.PropertyChanged += OnShellPropertyChanged;
    }

    public ObservableCollection<DkimSigningConfigDto> DkimConfigs { get; } = new();
    public ObservableCollection<HostedContentFilterPolicyDto> AntiSpamPolicies { get; } = new();
    public ObservableCollection<AntiPhishPolicyDto> AntiPhishPolicies { get; } = new();
    public ObservableCollection<MalwareFilterPolicyDto> MalwarePolicies { get; } = new();
    public ObservableCollection<QuarantinePolicyDto> QuarantinePolicies { get; } = new();
    public ObservableCollection<HostedOutboundSpamFilterPolicyDto> OutboundSpamPolicies { get; } = new();

    public IReadOnlyList<string> SpamActions { get; } = ["MoveToJmf", "AddXHeader", "ModifySubject", "Quarantine", "Delete"];
    public IReadOnlyList<string> AntiPhishActions { get; } = ["MoveToJmf", "Quarantine", "Delete"];
    public IReadOnlyList<string> MalwareFileTypeActions { get; } = ["Reject", "QuarantineMessage", "Quarantine"];
    public IReadOnlyList<string> QuarantinePermissions { get; } = ["FullAccess", "LimitedAccess", "Preview", "NoAccess"];
    public IReadOnlyList<string> OutboundThresholdActions { get; } = ["BlockUser", "NotifyOutboundSpam", "NotifyOutboundSpamAndBlockUser"];
    public IReadOnlyList<string> AutoForwardingModes { get; } = ["Automatic", "On", "Off"];

    public bool IsLoadingWorkspace
    {
        get => _isLoadingWorkspace;
        private set
        {
            if (SetProperty(ref _isLoadingWorkspace, value))
            {
                OnPropertyChanged(nameof(IsBusy));
                OnPropertyChanged(nameof(LoadingOverlayText));
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
                OnPropertyChanged(nameof(IsBusy));
                OnPropertyChanged(nameof(LoadingOverlayText));
                RaiseCanExecuteChanged();
            }
        }
    }

    public bool IsBusy => IsLoadingWorkspace || IsSaving;
    public string LoadingOverlayText => IsSaving
        ? "Saving Mail Security configuration..."
        : "Loading Mail Security workspace...";

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

    public string? WarningsText
    {
        get => _warningsText;
        private set
        {
            if (SetProperty(ref _warningsText, value))
            {
                OnPropertyChanged(nameof(HasWarnings));
            }
        }
    }

    public bool HasWarnings => !string.IsNullOrWhiteSpace(WarningsText);

    public DkimSigningConfigDto? SelectedDkimConfig
    {
        get => _selectedDkimConfig;
        set
        {
            if (SetProperty(ref _selectedDkimConfig, value))
            {
                DkimEnabled = value?.Enabled ?? false;
                RaiseCanExecuteChanged();
            }
        }
    }

    public HostedContentFilterPolicyDto? SelectedAntiSpamPolicy
    {
        get => _selectedAntiSpamPolicy;
        set
        {
            if (SetProperty(ref _selectedAntiSpamPolicy, value))
            {
                AntiSpamRuleEnabled = IsRuleEnabled(value?.RuleState);
                AntiSpamBulkThreshold = value?.BulkThreshold ?? 6;
                AntiSpamSpamAction = value?.SpamAction ?? "MoveToJmf";
                AntiSpamHighConfidenceSpamAction = value?.HighConfidenceSpamAction ?? "Quarantine";
                AntiSpamPhishSpamAction = value?.PhishSpamAction ?? "Quarantine";
                OnPropertyChanged(nameof(SelectedAntiSpamHasRule));
                RaiseCanExecuteChanged();
            }
        }
    }

    public AntiPhishPolicyDto? SelectedAntiPhishPolicy
    {
        get => _selectedAntiPhishPolicy;
        set
        {
            if (SetProperty(ref _selectedAntiPhishPolicy, value))
            {
                AntiPhishRuleEnabled = IsRuleEnabled(value?.RuleState);
                AntiPhishEnableSpoofIntelligence = value?.EnableSpoofIntelligence ?? false;
                AntiPhishEnableMailboxIntelligence = value?.EnableMailboxIntelligence ?? false;
                AntiPhishEnableTargetedUserProtection = value?.EnableTargetedUserProtection ?? false;
                AntiPhishHonorDmarcPolicy = value?.HonorDmarcPolicy ?? true;
                AntiPhishThresholdLevel = value?.PhishThresholdLevel ?? 2;
                AntiPhishMailboxIntelligenceProtectionAction = value?.MailboxIntelligenceProtectionAction ?? "Quarantine";
                AntiPhishTargetedUserProtectionAction = value?.TargetedUserProtectionAction ?? "Quarantine";
                AntiPhishAuthenticationFailAction = value?.AuthenticationFailAction ?? "MoveToJmf";
                AntiPhishDmarcRejectAction = value?.DmarcRejectAction ?? "Quarantine";
                AntiPhishDmarcQuarantineAction = value?.DmarcQuarantineAction ?? "Quarantine";
                OnPropertyChanged(nameof(SelectedAntiPhishHasRule));
                RaiseCanExecuteChanged();
            }
        }
    }

    public MalwareFilterPolicyDto? SelectedMalwarePolicy
    {
        get => _selectedMalwarePolicy;
        set
        {
            if (SetProperty(ref _selectedMalwarePolicy, value))
            {
                MalwareRuleEnabled = IsRuleEnabled(value?.RuleState);
                MalwareEnableFileFilter = value?.EnableFileFilter ?? false;
                MalwareZapEnabled = value?.ZapEnabled ?? true;
                MalwareFileTypeAction = value?.FileTypeAction ?? "Reject";
                OnPropertyChanged(nameof(SelectedMalwareHasRule));
                RaiseCanExecuteChanged();
            }
        }
    }

    public QuarantinePolicyDto? SelectedQuarantinePolicy
    {
        get => _selectedQuarantinePolicy;
        set
        {
            if (SetProperty(ref _selectedQuarantinePolicy, value))
            {
                QuarantinePermissionsValue = value?.EndUserQuarantinePermissionsValue ?? "FullAccess";
                QuarantineEsnEnabled = value?.EsnEnabled ?? true;
                QuarantineEndUserSpamNotificationFrequency = value?.EndUserSpamNotificationFrequency ?? 1;
                QuarantineRetentionDays = value?.QuarantineRetentionDays ?? 15;
                QuarantineOrganizationBrandingEnabled = value?.OrganizationBrandingEnabled ?? false;
                RaiseCanExecuteChanged();
            }
        }
    }

    public HostedOutboundSpamFilterPolicyDto? SelectedOutboundSpamPolicy
    {
        get => _selectedOutboundSpamPolicy;
        set
        {
            if (SetProperty(ref _selectedOutboundSpamPolicy, value))
            {
                OutboundRecipientLimitExternalPerHour = value?.RecipientLimitExternalPerHour ?? 500;
                OutboundRecipientLimitInternalPerHour = value?.RecipientLimitInternalPerHour ?? 1000;
                OutboundRecipientLimitPerDay = value?.RecipientLimitPerDay ?? 1000;
                OutboundActionWhenThresholdReached = value?.ActionWhenThresholdReached ?? "BlockUser";
                OutboundAutoForwardingMode = value?.AutoForwardingMode ?? "Automatic";
                RaiseCanExecuteChanged();
            }
        }
    }

    public bool DkimEnabled
    {
        get => _dkimEnabled;
        set => SetProperty(ref _dkimEnabled, value);
    }

    public bool AntiSpamRuleEnabled
    {
        get => _antiSpamRuleEnabled;
        set => SetProperty(ref _antiSpamRuleEnabled, value);
    }

    public int AntiSpamBulkThreshold
    {
        get => _antiSpamBulkThreshold;
        set => SetProperty(ref _antiSpamBulkThreshold, Math.Clamp(value, 1, 9));
    }

    public string AntiSpamSpamAction
    {
        get => _antiSpamSpamAction;
        set => SetProperty(ref _antiSpamSpamAction, value);
    }

    public string AntiSpamHighConfidenceSpamAction
    {
        get => _antiSpamHighConfidenceSpamAction;
        set => SetProperty(ref _antiSpamHighConfidenceSpamAction, value);
    }

    public string AntiSpamPhishSpamAction
    {
        get => _antiSpamPhishSpamAction;
        set => SetProperty(ref _antiSpamPhishSpamAction, value);
    }

    public bool AntiPhishRuleEnabled
    {
        get => _antiPhishRuleEnabled;
        set => SetProperty(ref _antiPhishRuleEnabled, value);
    }

    public bool AntiPhishEnableSpoofIntelligence
    {
        get => _antiPhishEnableSpoofIntelligence;
        set => SetProperty(ref _antiPhishEnableSpoofIntelligence, value);
    }

    public bool AntiPhishEnableMailboxIntelligence
    {
        get => _antiPhishEnableMailboxIntelligence;
        set => SetProperty(ref _antiPhishEnableMailboxIntelligence, value);
    }

    public bool AntiPhishEnableTargetedUserProtection
    {
        get => _antiPhishEnableTargetedUserProtection;
        set => SetProperty(ref _antiPhishEnableTargetedUserProtection, value);
    }

    public bool AntiPhishHonorDmarcPolicy
    {
        get => _antiPhishHonorDmarcPolicy;
        set => SetProperty(ref _antiPhishHonorDmarcPolicy, value);
    }

    public int AntiPhishThresholdLevel
    {
        get => _antiPhishThresholdLevel;
        set => SetProperty(ref _antiPhishThresholdLevel, Math.Clamp(value, 1, 4));
    }

    public string AntiPhishMailboxIntelligenceProtectionAction
    {
        get => _antiPhishMailboxIntelligenceProtectionAction;
        set => SetProperty(ref _antiPhishMailboxIntelligenceProtectionAction, value);
    }

    public string AntiPhishTargetedUserProtectionAction
    {
        get => _antiPhishTargetedUserProtectionAction;
        set => SetProperty(ref _antiPhishTargetedUserProtectionAction, value);
    }

    public string AntiPhishAuthenticationFailAction
    {
        get => _antiPhishAuthenticationFailAction;
        set => SetProperty(ref _antiPhishAuthenticationFailAction, value);
    }

    public string AntiPhishDmarcRejectAction
    {
        get => _antiPhishDmarcRejectAction;
        set => SetProperty(ref _antiPhishDmarcRejectAction, value);
    }

    public string AntiPhishDmarcQuarantineAction
    {
        get => _antiPhishDmarcQuarantineAction;
        set => SetProperty(ref _antiPhishDmarcQuarantineAction, value);
    }

    public bool MalwareRuleEnabled
    {
        get => _malwareRuleEnabled;
        set => SetProperty(ref _malwareRuleEnabled, value);
    }

    public bool MalwareEnableFileFilter
    {
        get => _malwareEnableFileFilter;
        set => SetProperty(ref _malwareEnableFileFilter, value);
    }

    public bool MalwareZapEnabled
    {
        get => _malwareZapEnabled;
        set => SetProperty(ref _malwareZapEnabled, value);
    }

    public string MalwareFileTypeAction
    {
        get => _malwareFileTypeAction;
        set => SetProperty(ref _malwareFileTypeAction, value);
    }

    public string QuarantinePermissionsValue
    {
        get => _quarantinePermissions;
        set => SetProperty(ref _quarantinePermissions, value);
    }

    public bool QuarantineEsnEnabled
    {
        get => _quarantineEsnEnabled;
        set => SetProperty(ref _quarantineEsnEnabled, value);
    }

    public int QuarantineEndUserSpamNotificationFrequency
    {
        get => _quarantineEndUserSpamNotificationFrequency;
        set => SetProperty(ref _quarantineEndUserSpamNotificationFrequency, Math.Clamp(value, 1, 15));
    }

    public int QuarantineRetentionDays
    {
        get => _quarantineRetentionDays;
        set => SetProperty(ref _quarantineRetentionDays, Math.Clamp(value, 1, 30));
    }

    public bool QuarantineOrganizationBrandingEnabled
    {
        get => _quarantineOrganizationBrandingEnabled;
        set => SetProperty(ref _quarantineOrganizationBrandingEnabled, value);
    }

    public int OutboundRecipientLimitExternalPerHour
    {
        get => _outboundRecipientLimitExternalPerHour;
        set => SetProperty(ref _outboundRecipientLimitExternalPerHour, Math.Max(value, 0));
    }

    public int OutboundRecipientLimitInternalPerHour
    {
        get => _outboundRecipientLimitInternalPerHour;
        set => SetProperty(ref _outboundRecipientLimitInternalPerHour, Math.Max(value, 0));
    }

    public int OutboundRecipientLimitPerDay
    {
        get => _outboundRecipientLimitPerDay;
        set => SetProperty(ref _outboundRecipientLimitPerDay, Math.Max(value, 0));
    }

    public string OutboundActionWhenThresholdReached
    {
        get => _outboundActionWhenThresholdReached;
        set => SetProperty(ref _outboundActionWhenThresholdReached, value);
    }

    public string OutboundAutoForwardingMode
    {
        get => _outboundAutoForwardingMode;
        set => SetProperty(ref _outboundAutoForwardingMode, value);
    }

    public bool SelectedAntiSpamHasRule => !string.IsNullOrWhiteSpace(SelectedAntiSpamPolicy?.RuleIdentity);
    public bool SelectedAntiPhishHasRule => !string.IsNullOrWhiteSpace(SelectedAntiPhishPolicy?.RuleIdentity);
    public bool SelectedMalwareHasRule => !string.IsNullOrWhiteSpace(SelectedMalwarePolicy?.RuleIdentity);

    public bool CanRefreshWorkspace => _shellViewModel.IsExchangeConnected && !IsBusy;
    public bool CanSaveDkim => _shellViewModel.IsExchangeConnected && !IsBusy && SelectedDkimConfig != null;
    public bool CanSaveAntiSpam => _shellViewModel.IsExchangeConnected && !IsBusy && SelectedAntiSpamPolicy != null;
    public bool CanSaveAntiPhish => _shellViewModel.IsExchangeConnected && !IsBusy && SelectedAntiPhishPolicy != null;
    public bool CanSaveMalware => _shellViewModel.IsExchangeConnected && !IsBusy && SelectedMalwarePolicy != null;
    public bool CanSaveQuarantine => _shellViewModel.IsExchangeConnected && !IsBusy && SelectedQuarantinePolicy != null;
    public bool CanSaveOutboundSpam => _shellViewModel.IsExchangeConnected && !IsBusy && SelectedOutboundSpamPolicy != null;

    public ICommand RefreshWorkspaceCommand { get; }
    public ICommand SaveDkimCommand { get; }
    public ICommand SaveAntiSpamCommand { get; }
    public ICommand SaveAntiPhishCommand { get; }
    public ICommand SaveMalwareCommand { get; }
    public ICommand SaveQuarantineCommand { get; }
    public ICommand SaveOutboundSpamCommand { get; }

    public async Task LoadAsync()
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            ClearStateForDisconnectedSession();
            return;
        }

        if (DkimConfigs.Count == 0 &&
            AntiSpamPolicies.Count == 0 &&
            AntiPhishPolicies.Count == 0 &&
            MalwarePolicies.Count == 0 &&
            QuarantinePolicies.Count == 0 &&
            OutboundSpamPolicies.Count == 0)
        {
            await RefreshWorkspaceAsync(CancellationToken.None);
        }
    }

    private async Task RefreshWorkspaceAsync(CancellationToken cancellationToken)
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            ClearStateForDisconnectedSession();
            return;
        }

        IsLoadingWorkspace = true;
        ErrorMessage = null;
        WarningsText = null;

        var selectedDkimIdentity = SelectedDkimConfig?.Identity;
        var selectedAntiSpamIdentity = SelectedAntiSpamPolicy?.Identity;
        var selectedAntiPhishIdentity = SelectedAntiPhishPolicy?.Identity;
        var selectedMalwareIdentity = SelectedMalwarePolicy?.Identity;
        var selectedQuarantineIdentity = SelectedQuarantinePolicy?.Identity;
        var selectedOutboundIdentity = SelectedOutboundSpamPolicy?.Identity;

        try
        {
            var result = await _workerService.GetMailSecurityBaselineAsync(cancellationToken: cancellationToken);
            if (!result.IsSuccess || result.Value == null)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to load the Mail Security workspace.";
                return;
            }

            DkimConfigs.ReplaceAll(result.Value.DkimConfigs);
            AntiSpamPolicies.ReplaceAll(result.Value.AntiSpamPolicies);
            AntiPhishPolicies.ReplaceAll(result.Value.AntiPhishPolicies);
            MalwarePolicies.ReplaceAll(result.Value.MalwarePolicies);
            QuarantinePolicies.ReplaceAll(result.Value.QuarantinePolicies);
            OutboundSpamPolicies.ReplaceAll(result.Value.OutboundSpamPolicies);

            WarningsText = result.Value.Warnings.Count == 0
                ? null
                : string.Join(Environment.NewLine, result.Value.Warnings);

            SelectedDkimConfig = DkimConfigs.FirstOrDefault(item => string.Equals(item.Identity, selectedDkimIdentity, StringComparison.OrdinalIgnoreCase)) ?? DkimConfigs.FirstOrDefault();
            SelectedAntiSpamPolicy = AntiSpamPolicies.FirstOrDefault(item => string.Equals(item.Identity, selectedAntiSpamIdentity, StringComparison.OrdinalIgnoreCase)) ?? AntiSpamPolicies.FirstOrDefault();
            SelectedAntiPhishPolicy = AntiPhishPolicies.FirstOrDefault(item => string.Equals(item.Identity, selectedAntiPhishIdentity, StringComparison.OrdinalIgnoreCase)) ?? AntiPhishPolicies.FirstOrDefault();
            SelectedMalwarePolicy = MalwarePolicies.FirstOrDefault(item => string.Equals(item.Identity, selectedMalwareIdentity, StringComparison.OrdinalIgnoreCase)) ?? MalwarePolicies.FirstOrDefault();
            SelectedQuarantinePolicy = QuarantinePolicies.FirstOrDefault(item => string.Equals(item.Identity, selectedQuarantineIdentity, StringComparison.OrdinalIgnoreCase)) ?? QuarantinePolicies.FirstOrDefault();
            SelectedOutboundSpamPolicy = OutboundSpamPolicies.FirstOrDefault(item => string.Equals(item.Identity, selectedOutboundIdentity, StringComparison.OrdinalIgnoreCase)) ?? OutboundSpamPolicies.FirstOrDefault();
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsLoadingWorkspace = false;
        }
    }

    private async Task SaveDkimAsync(CancellationToken cancellationToken)
    {
        if (SelectedDkimConfig == null)
        {
            return;
        }

        await ExecuteSaveAsync(
            "Updating DKIM",
            SelectedDkimConfig.Domain,
            "Enable or disable DKIM signing for the selected domain.",
            () => _workerService.UpdateDkimSigningConfigAsync(
                new UpdateDkimSigningConfigRequest
                {
                    Identity = SelectedDkimConfig.Identity,
                    Enabled = DkimEnabled
                },
                cancellationToken: cancellationToken),
            $"DKIM updated for {SelectedDkimConfig.Domain}",
            cancellationToken);
    }

    private async Task SaveAntiSpamAsync(CancellationToken cancellationToken)
    {
        if (SelectedAntiSpamPolicy == null)
        {
            return;
        }

        await ExecuteSaveAsync(
            "Updating anti-spam policy",
            SelectedAntiSpamPolicy.Name,
            "Update the thresholds and actions for the selected anti-spam policy.",
            () => _workerService.UpdateHostedContentFilterPolicyAsync(
                new UpdateHostedContentFilterPolicyRequest
                {
                    Identity = SelectedAntiSpamPolicy.Identity,
                    RuleIdentity = SelectedAntiSpamPolicy.RuleIdentity,
                    Enabled = SelectedAntiSpamHasRule ? AntiSpamRuleEnabled : null,
                    BulkThreshold = AntiSpamBulkThreshold,
                    SpamAction = TrimToNull(AntiSpamSpamAction),
                    HighConfidenceSpamAction = TrimToNull(AntiSpamHighConfidenceSpamAction),
                    PhishSpamAction = TrimToNull(AntiSpamPhishSpamAction)
                },
                cancellationToken: cancellationToken),
            $"Anti-spam policy updated: {SelectedAntiSpamPolicy.Name}",
            cancellationToken);
    }

    private async Task SaveAntiPhishAsync(CancellationToken cancellationToken)
    {
        if (SelectedAntiPhishPolicy == null)
        {
            return;
        }

        await ExecuteSaveAsync(
            "Updating anti-phish policy",
            SelectedAntiPhishPolicy.Name,
            "Update anti-phish and spoof protection for the selected policy.",
            () => _workerService.UpdateAntiPhishPolicyAsync(
                new UpdateAntiPhishPolicyRequest
                {
                    Identity = SelectedAntiPhishPolicy.Identity,
                    RuleIdentity = SelectedAntiPhishPolicy.RuleIdentity,
                    Enabled = SelectedAntiPhishHasRule ? AntiPhishRuleEnabled : null,
                    EnableSpoofIntelligence = AntiPhishEnableSpoofIntelligence,
                    EnableMailboxIntelligence = AntiPhishEnableMailboxIntelligence,
                    EnableTargetedUserProtection = AntiPhishEnableTargetedUserProtection,
                    HonorDmarcPolicy = AntiPhishHonorDmarcPolicy,
                    PhishThresholdLevel = AntiPhishThresholdLevel,
                    MailboxIntelligenceProtectionAction = TrimToNull(AntiPhishMailboxIntelligenceProtectionAction),
                    TargetedUserProtectionAction = TrimToNull(AntiPhishTargetedUserProtectionAction),
                    AuthenticationFailAction = TrimToNull(AntiPhishAuthenticationFailAction),
                    DmarcRejectAction = TrimToNull(AntiPhishDmarcRejectAction),
                    DmarcQuarantineAction = TrimToNull(AntiPhishDmarcQuarantineAction)
                },
                cancellationToken: cancellationToken),
            $"Anti-phish policy updated: {SelectedAntiPhishPolicy.Name}",
            cancellationToken);
    }

    private async Task SaveMalwareAsync(CancellationToken cancellationToken)
    {
        if (SelectedMalwarePolicy == null)
        {
            return;
        }

        await ExecuteSaveAsync(
            "Updating anti-malware policy",
            SelectedMalwarePolicy.Name,
            "Update the file filter and anti-malware actions for the selected policy.",
            () => _workerService.UpdateMalwareFilterPolicyAsync(
                new UpdateMalwareFilterPolicyRequest
                {
                    Identity = SelectedMalwarePolicy.Identity,
                    RuleIdentity = SelectedMalwarePolicy.RuleIdentity,
                    Enabled = SelectedMalwareHasRule ? MalwareRuleEnabled : null,
                    EnableFileFilter = MalwareEnableFileFilter,
                    ZapEnabled = MalwareZapEnabled,
                    FileTypeAction = TrimToNull(MalwareFileTypeAction)
                },
                cancellationToken: cancellationToken),
            $"Anti-malware policy updated: {SelectedMalwarePolicy.Name}",
            cancellationToken);
    }

    private async Task SaveQuarantineAsync(CancellationToken cancellationToken)
    {
        if (SelectedQuarantinePolicy == null)
        {
            return;
        }

        await ExecuteSaveAsync(
            "Updating quarantine policy",
            SelectedQuarantinePolicy.Name,
            "Update retention and end-user permissions for the selected quarantine policy.",
            () => _workerService.UpdateQuarantinePolicyAsync(
                new UpdateQuarantinePolicyRequest
                {
                    Identity = SelectedQuarantinePolicy.Identity,
                    EndUserQuarantinePermissionsValue = TrimToNull(QuarantinePermissionsValue),
                    EsnEnabled = QuarantineEsnEnabled,
                    EndUserSpamNotificationFrequency = QuarantineEndUserSpamNotificationFrequency,
                    QuarantineRetentionDays = QuarantineRetentionDays,
                    OrganizationBrandingEnabled = QuarantineOrganizationBrandingEnabled
                },
                cancellationToken: cancellationToken),
            $"Quarantine policy updated: {SelectedQuarantinePolicy.Name}",
            cancellationToken);
    }

    private async Task SaveOutboundSpamAsync(CancellationToken cancellationToken)
    {
        if (SelectedOutboundSpamPolicy == null)
        {
            return;
        }

        await ExecuteSaveAsync(
            "Updating outbound spam policy",
            SelectedOutboundSpamPolicy.Name,
            "Update the limits and actions for the selected outbound spam policy.",
            () => _workerService.UpdateHostedOutboundSpamFilterPolicyAsync(
                new UpdateHostedOutboundSpamFilterPolicyRequest
                {
                    Identity = SelectedOutboundSpamPolicy.Identity,
                    RecipientLimitExternalPerHour = OutboundRecipientLimitExternalPerHour,
                    RecipientLimitInternalPerHour = OutboundRecipientLimitInternalPerHour,
                    RecipientLimitPerDay = OutboundRecipientLimitPerDay,
                    ActionWhenThresholdReached = TrimToNull(OutboundActionWhenThresholdReached),
                    AutoForwardingMode = TrimToNull(OutboundAutoForwardingMode)
                },
                cancellationToken: cancellationToken),
            $"Outbound spam policy updated: {SelectedOutboundSpamPolicy.Name}",
            cancellationToken);
    }

    private async Task ExecuteSaveAsync(
        string operation,
        string target,
        string impact,
        Func<Task<Result>> saveAction,
        string logMessage,
        CancellationToken cancellationToken)
    {
        if (!ConfirmMutation(operation, target, impact, "Confirm Mail Security update"))
        {
            return;
        }

        IsSaving = true;
        ErrorMessage = null;

        try
        {
            var result = await saveAction();
            if (!result.IsSuccess)
            {
                ErrorMessage = result.Error?.Message ?? "Updating Mail Security failed.";
                return;
            }

            _shellViewModel.AddLog(LogLevel.Information, logMessage, "MailSecurity");
            await RefreshWorkspaceAsync(cancellationToken);
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

    private void OnShellPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ShellViewModel.IsExchangeConnected))
        {
            if (!_shellViewModel.IsExchangeConnected)
            {
                ClearStateForDisconnectedSession();
            }

            RaiseCanExecuteChanged();
        }
    }

    private void ClearStateForDisconnectedSession()
    {
        DkimConfigs.Clear();
        AntiSpamPolicies.Clear();
        AntiPhishPolicies.Clear();
        MalwarePolicies.Clear();
        QuarantinePolicies.Clear();
        OutboundSpamPolicies.Clear();
        SelectedDkimConfig = null;
        SelectedAntiSpamPolicy = null;
        SelectedAntiPhishPolicy = null;
        SelectedMalwarePolicy = null;
        SelectedQuarantinePolicy = null;
        SelectedOutboundSpamPolicy = null;
        WarningsText = null;
        ErrorMessage = "Not connected to Exchange Online";
    }

    private void RaiseCanExecuteChanged()
    {
        OnPropertyChanged(nameof(CanRefreshWorkspace));
        OnPropertyChanged(nameof(CanSaveDkim));
        OnPropertyChanged(nameof(CanSaveAntiSpam));
        OnPropertyChanged(nameof(CanSaveAntiPhish));
        OnPropertyChanged(nameof(CanSaveMalware));
        OnPropertyChanged(nameof(CanSaveQuarantine));
        OnPropertyChanged(nameof(CanSaveOutboundSpam));
        CommandManager.InvalidateRequerySuggested();
    }

    private static bool IsRuleEnabled(string? ruleState)
        => string.Equals(ruleState, "Enabled", StringComparison.OrdinalIgnoreCase);

    private static string? TrimToNull(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : value.Trim();
}
