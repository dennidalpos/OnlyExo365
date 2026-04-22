using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Input;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;

namespace OnlyExo365.Shell.ViewModels;

public sealed class MailboxSettingsEditorViewModel : ViewModelBase
{
    private static readonly TimeSpan RetentionPolicyCacheTtl = TimeSpan.FromMinutes(10);

    private readonly IMailboxesWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;
    private readonly CacheService _cacheService;
    private MailboxCasCapabilityState _casCapabilityState = MailboxCasCapabilityState.Unknown();

    private bool _hasPendingChanges;
    private bool _isInitializing;
    private bool _isRefreshingRetentionPolicies;
    private bool _isRetentionPolicyLoading;
    private MailboxSettingsSnapshot? _originalSettings;
    private string? _retentionPolicyFallback;
    private string? _primarySmtpAddress;
    private string? _proxyAddressesText;
    private bool _hiddenFromAddressListsEnabled;
    private string? _forwardingAddress;
    private string? _forwardingSmtpAddress;
    private bool _deliverToMailboxAndForward;
    private bool _archiveEnabled;
    private bool _litigationHoldEnabled;
    private bool _auditEnabled;
    private bool _singleItemRecoveryEnabled;
    private bool _retentionHoldEnabled;
    private string? _issueWarningQuota;
    private string? _prohibitSendQuota;
    private string? _prohibitSendReceiveQuota;
    private string? _maxSendSize;
    private string? _maxReceiveSize;
    private string? _selectedRetentionPolicy;
    private bool _owaEnabled;
    private bool _activeSyncEnabled;
    private bool _mapiEnabled;
    private bool _popEnabled;
    private bool _imapEnabled;
    private bool _smtpClientAuthenticationDisabled;

    public MailboxSettingsEditorViewModel(IMailboxesWorkerService workerService, ShellViewModel shellViewModel, CacheService cacheService)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;
        _cacheService = cacheService;
        _shellViewModel.PropertyChanged += OnShellViewModelPropertyChanged;
        ApplyCapabilities(_shellViewModel.Capabilities);
    }

    public ObservableCollection<RetentionPolicySummaryDto> AvailableRetentionPolicies { get; } = new();

    public bool HasPendingChanges
    {
        get => _hasPendingChanges;
        private set => SetProperty(ref _hasPendingChanges, value);
    }

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

    public string? PrimarySmtpAddress
    {
        get => _primarySmtpAddress;
        set
        {
            if (SetProperty(ref _primarySmtpAddress, value))
            {
                UpdatePendingChanges();
            }
        }
    }

    public string? ProxyAddressesText
    {
        get => _proxyAddressesText;
        set
        {
            if (SetProperty(ref _proxyAddressesText, value))
            {
                UpdatePendingChanges();
            }
        }
    }

    public bool HiddenFromAddressListsEnabled
    {
        get => _hiddenFromAddressListsEnabled;
        set
        {
            if (SetProperty(ref _hiddenFromAddressListsEnabled, value))
            {
                UpdatePendingChanges();
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
                UpdatePendingChanges();
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
                UpdatePendingChanges();
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
                UpdatePendingChanges();
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
                UpdatePendingChanges();
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
                UpdatePendingChanges();
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
                UpdatePendingChanges();
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
                UpdatePendingChanges();
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
                UpdatePendingChanges();
            }
        }
    }

    public string? IssueWarningQuota
    {
        get => _issueWarningQuota;
        set
        {
            if (SetProperty(ref _issueWarningQuota, value))
            {
                UpdatePendingChanges();
            }
        }
    }

    public string? ProhibitSendQuota
    {
        get => _prohibitSendQuota;
        set
        {
            if (SetProperty(ref _prohibitSendQuota, value))
            {
                UpdatePendingChanges();
            }
        }
    }

    public string? ProhibitSendReceiveQuota
    {
        get => _prohibitSendReceiveQuota;
        set
        {
            if (SetProperty(ref _prohibitSendReceiveQuota, value))
            {
                UpdatePendingChanges();
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
                UpdatePendingChanges();
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
                UpdatePendingChanges();
            }
        }
    }

    public string? SelectedRetentionPolicy
    {
        get => _selectedRetentionPolicy;
        set
        {
            if (SetProperty(ref _selectedRetentionPolicy, value))
            {
                UpdatePendingChanges();
                OnPropertyChanged(nameof(SelectedRetentionPolicyDescription));
                OnPropertyChanged(nameof(SelectedRetentionPolicyRequiresArchive));
                OnPropertyChanged(nameof(ShowArchiveRequiredWarning));
            }
        }
    }

    public string? SelectedRetentionPolicyDescription
        => AvailableRetentionPolicies.FirstOrDefault(IsPolicySelected)?.Description;

    public bool SelectedRetentionPolicyRequiresArchive
        => AvailableRetentionPolicies.FirstOrDefault(IsPolicySelected)?.RequiresArchive ?? false;

    public bool ShowArchiveRequiredWarning => !ArchiveEnabled && SelectedRetentionPolicyRequiresArchive;

    public bool OwaEnabled
    {
        get => _owaEnabled;
        set
        {
            if (SetProperty(ref _owaEnabled, value))
            {
                UpdatePendingChanges();
            }
        }
    }

    public bool ActiveSyncEnabled
    {
        get => _activeSyncEnabled;
        set
        {
            if (SetProperty(ref _activeSyncEnabled, value))
            {
                UpdatePendingChanges();
            }
        }
    }

    public bool MapiEnabled
    {
        get => _mapiEnabled;
        set
        {
            if (SetProperty(ref _mapiEnabled, value))
            {
                UpdatePendingChanges();
            }
        }
    }

    public bool PopEnabled
    {
        get => _popEnabled;
        set
        {
            if (SetProperty(ref _popEnabled, value))
            {
                UpdatePendingChanges();
            }
        }
    }

    public bool ImapEnabled
    {
        get => _imapEnabled;
        set
        {
            if (SetProperty(ref _imapEnabled, value))
            {
                UpdatePendingChanges();
            }
        }
    }

    public bool SmtpClientAuthenticationDisabled
    {
        get => _smtpClientAuthenticationDisabled;
        set
        {
            if (SetProperty(ref _smtpClientAuthenticationDisabled, value))
            {
                UpdatePendingChanges();
            }
        }
    }

    public bool CanEditOwaEnabled => _casCapabilityState.CanReadSettings && _casCapabilityState.CanEditOwaEnabled;

    public bool CanEditActiveSyncEnabled => _casCapabilityState.CanReadSettings && _casCapabilityState.CanEditActiveSyncEnabled;

    public bool CanEditMapiEnabled => _casCapabilityState.CanReadSettings && _casCapabilityState.CanEditMapiEnabled;

    public bool CanEditPopEnabled => _casCapabilityState.CanReadSettings && _casCapabilityState.CanEditPopEnabled;

    public bool CanEditImapEnabled => _casCapabilityState.CanReadSettings && _casCapabilityState.CanEditImapEnabled;

    public bool CanEditSmtpClientAuthenticationDisabled =>
        _casCapabilityState.CanReadSettings && _casCapabilityState.CanEditSmtpClientAuthenticationDisabled;

    public string? CasCapabilityMessage => _casCapabilityState.Message;

    public void Initialize(MailboxDetailsDto? details)
    {
        if (details?.Features == null)
        {
            Reset();
            return;
        }

        _isInitializing = true;

        var features = details.Features;
        _retentionPolicyFallback = !string.IsNullOrWhiteSpace(details.RetentionPolicy)
            ? details.RetentionPolicy
            : _retentionPolicyFallback;

        PrimarySmtpAddress = details.PrimarySmtpAddress ?? string.Empty;
        ProxyAddressesText = BuildProxyAddressesText(details.EmailAddresses, details.PrimarySmtpAddress);
        HiddenFromAddressListsEnabled = features.HiddenFromAddressListsEnabled;
        ForwardingAddress = features.ForwardingAddress ?? string.Empty;
        ForwardingSmtpAddress = features.ForwardingSmtpAddress ?? string.Empty;
        DeliverToMailboxAndForward = features.DeliverToMailboxAndForward;
        ArchiveEnabled = features.ArchiveEnabled;
        LitigationHoldEnabled = features.LitigationHoldEnabled;
        AuditEnabled = features.AuditEnabled;
        SingleItemRecoveryEnabled = features.SingleItemRecoveryEnabled;
        RetentionHoldEnabled = features.RetentionHoldEnabled;
        IssueWarningQuota = features.IssueWarningQuota ?? string.Empty;
        ProhibitSendQuota = features.ProhibitSendQuota ?? string.Empty;
        ProhibitSendReceiveQuota = features.ProhibitSendReceiveQuota ?? string.Empty;
        MaxSendSize = features.MaxSendSize ?? string.Empty;
        MaxReceiveSize = features.MaxReceiveSize ?? string.Empty;
        OwaEnabled = features.OwaEnabled ?? false;
        ActiveSyncEnabled = features.ActiveSyncEnabled ?? false;
        MapiEnabled = features.MapiEnabled ?? false;
        PopEnabled = features.PopEnabled ?? false;
        ImapEnabled = features.ImapEnabled ?? false;
        SmtpClientAuthenticationDisabled = features.SmtpClientAuthenticationDisabled ?? false;
        SelectedRetentionPolicy = _retentionPolicyFallback ?? string.Empty;

        _originalSettings = CaptureSnapshot();
        _isInitializing = false;
        HasPendingChanges = false;
    }

    public void Reset()
    {
        _isInitializing = true;
        _retentionPolicyFallback = null;
        PrimarySmtpAddress = string.Empty;
        ProxyAddressesText = string.Empty;
        HiddenFromAddressListsEnabled = false;
        ForwardingAddress = string.Empty;
        ForwardingSmtpAddress = string.Empty;
        DeliverToMailboxAndForward = false;
        ArchiveEnabled = false;
        LitigationHoldEnabled = false;
        AuditEnabled = false;
        SingleItemRecoveryEnabled = false;
        RetentionHoldEnabled = false;
        IssueWarningQuota = string.Empty;
        ProhibitSendQuota = string.Empty;
        ProhibitSendReceiveQuota = string.Empty;
        MaxSendSize = string.Empty;
        MaxReceiveSize = string.Empty;
        OwaEnabled = false;
        ActiveSyncEnabled = false;
        MapiEnabled = false;
        PopEnabled = false;
        ImapEnabled = false;
        SmtpClientAuthenticationDisabled = false;
        SelectedRetentionPolicy = string.Empty;
        _originalSettings = null;
        _isInitializing = false;
        HasPendingChanges = false;
    }

    public void ApplyCapabilities(CapabilityMapDto? capabilities)
    {
        _casCapabilityState = MailboxCasCapabilityState.From(capabilities);
        OnPropertyChanged(nameof(CanEditOwaEnabled));
        OnPropertyChanged(nameof(CanEditActiveSyncEnabled));
        OnPropertyChanged(nameof(CanEditMapiEnabled));
        OnPropertyChanged(nameof(CanEditPopEnabled));
        OnPropertyChanged(nameof(CanEditImapEnabled));
        OnPropertyChanged(nameof(CanEditSmtpClientAuthenticationDisabled));
        OnPropertyChanged(nameof(CasCapabilityMessage));
    }

    public async Task LoadRetentionPoliciesAsync(CancellationToken cancellationToken)
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
                    _cacheService.Set(CacheService.Keys.RetentionPolicies, policies, RetentionPolicyCacheTtl);
                }
                else if (!result.WasCancelled)
                {
                    _shellViewModel.AddLog(LogLevel.Warning, result.Error?.Message ?? "Unable to retrieve retention policies.");
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
                        Description = "No policy assigned"
                    });

                    foreach (var policy in policies)
                    {
                        AvailableRetentionPolicies.Add(policy);
                    }

                    if (selectedPolicySnapshot != null)
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
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Retention policies load failed: {ex.Message}");
        }
        finally
        {
            _isRefreshingRetentionPolicies = false;
            UpdatePendingChanges();
            IsRetentionPolicyLoading = false;
        }
    }

    public UpdateMailboxSettingsRequest BuildUpdateRequest(
        string identity,
        out bool settingsChanged,
        out bool retentionPolicyChanged,
        out string? retentionPolicyOverride)
    {
        var request = new UpdateMailboxSettingsRequest { Identity = identity };
        settingsChanged = false;

        if (_originalSettings == null)
        {
            retentionPolicyChanged = false;
            retentionPolicyOverride = null;
            return request;
        }

        var normalizedForwardingAddress = NormalizeInput(ForwardingAddress);
        var normalizedForwardingSmtp = NormalizeInput(ForwardingSmtpAddress);
        var normalizedPrimarySmtpAddress = NormalizeInput(PrimarySmtpAddress);
        var normalizedProxyAddresses = ParseProxyAddresses(ProxyAddressesText);
        var normalizedIssueWarningQuota = NormalizeInput(IssueWarningQuota);
        var normalizedProhibitSendQuota = NormalizeInput(ProhibitSendQuota);
        var normalizedProhibitSendReceiveQuota = NormalizeInput(ProhibitSendReceiveQuota);
        var normalizedMaxSend = NormalizeInput(MaxSendSize);
        var normalizedMaxReceive = NormalizeInput(MaxReceiveSize);
        retentionPolicyOverride = NormalizeInput(SelectedRetentionPolicy);

        var smtpConfigurationChanged = !string.Equals(normalizedPrimarySmtpAddress, _originalSettings.PrimarySmtpAddress, StringComparison.Ordinal)
            || !normalizedProxyAddresses.SequenceEqual(_originalSettings.ProxyAddresses, StringComparer.Ordinal);

        if (smtpConfigurationChanged)
        {
            if (string.IsNullOrWhiteSpace(normalizedPrimarySmtpAddress))
            {
                throw new InvalidOperationException("Primary SMTP is required when updating mailbox addresses.");
            }

            request.PrimarySmtpAddress = normalizedPrimarySmtpAddress;
            request.ProxyAddresses = normalizedProxyAddresses;
            settingsChanged = true;
        }

        if (HiddenFromAddressListsEnabled != _originalSettings.HiddenFromAddressListsEnabled)
        {
            request.HiddenFromAddressListsEnabled = HiddenFromAddressListsEnabled;
            settingsChanged = true;
        }

        if (ArchiveEnabled != _originalSettings.ArchiveEnabled)
        {
            request.ArchiveEnabled = ArchiveEnabled;
            settingsChanged = true;
        }

        if (LitigationHoldEnabled != _originalSettings.LitigationHoldEnabled)
        {
            request.LitigationHoldEnabled = LitigationHoldEnabled;
            settingsChanged = true;
        }

        if (AuditEnabled != _originalSettings.AuditEnabled)
        {
            request.AuditEnabled = AuditEnabled;
            settingsChanged = true;
        }

        if (SingleItemRecoveryEnabled != _originalSettings.SingleItemRecoveryEnabled)
        {
            request.SingleItemRecoveryEnabled = SingleItemRecoveryEnabled;
            settingsChanged = true;
        }

        if (RetentionHoldEnabled != _originalSettings.RetentionHoldEnabled)
        {
            request.RetentionHoldEnabled = RetentionHoldEnabled;
            settingsChanged = true;
        }

        if (!string.Equals(normalizedIssueWarningQuota, _originalSettings.IssueWarningQuota, StringComparison.Ordinal))
        {
            request.IssueWarningQuota = IssueWarningQuota ?? string.Empty;
            settingsChanged = true;
        }

        if (!string.Equals(normalizedProhibitSendQuota, _originalSettings.ProhibitSendQuota, StringComparison.Ordinal))
        {
            request.ProhibitSendQuota = ProhibitSendQuota ?? string.Empty;
            settingsChanged = true;
        }

        if (!string.Equals(normalizedProhibitSendReceiveQuota, _originalSettings.ProhibitSendReceiveQuota, StringComparison.Ordinal))
        {
            request.ProhibitSendReceiveQuota = ProhibitSendReceiveQuota ?? string.Empty;
            settingsChanged = true;
        }

        if (!string.Equals(normalizedForwardingAddress, _originalSettings.ForwardingAddress, StringComparison.Ordinal))
        {
            request.ForwardingAddress = ForwardingAddress ?? string.Empty;
            settingsChanged = true;
        }

        if (!string.Equals(normalizedForwardingSmtp, _originalSettings.ForwardingSmtpAddress, StringComparison.Ordinal))
        {
            request.ForwardingSmtpAddress = ForwardingSmtpAddress ?? string.Empty;
            settingsChanged = true;
        }

        if (DeliverToMailboxAndForward != _originalSettings.DeliverToMailboxAndForward)
        {
            request.DeliverToMailboxAndForward = DeliverToMailboxAndForward;
            settingsChanged = true;
        }

        if (!string.Equals(normalizedMaxSend, _originalSettings.MaxSendSize, StringComparison.Ordinal))
        {
            request.MaxSendSize = MaxSendSize ?? string.Empty;
            settingsChanged = true;
        }

        if (!string.Equals(normalizedMaxReceive, _originalSettings.MaxReceiveSize, StringComparison.Ordinal))
        {
            request.MaxReceiveSize = MaxReceiveSize ?? string.Empty;
            settingsChanged = true;
        }

        if (OwaEnabled != _originalSettings.OwaEnabled)
        {
            request.OwaEnabled = OwaEnabled;
            settingsChanged = true;
        }

        if (ActiveSyncEnabled != _originalSettings.ActiveSyncEnabled)
        {
            request.ActiveSyncEnabled = ActiveSyncEnabled;
            settingsChanged = true;
        }

        if (MapiEnabled != _originalSettings.MapiEnabled)
        {
            request.MapiEnabled = MapiEnabled;
            settingsChanged = true;
        }

        if (PopEnabled != _originalSettings.PopEnabled)
        {
            request.PopEnabled = PopEnabled;
            settingsChanged = true;
        }

        if (ImapEnabled != _originalSettings.ImapEnabled)
        {
            request.ImapEnabled = ImapEnabled;
            settingsChanged = true;
        }

        if (SmtpClientAuthenticationDisabled != _originalSettings.SmtpClientAuthenticationDisabled)
        {
            request.SmtpClientAuthenticationDisabled = SmtpClientAuthenticationDisabled;
            settingsChanged = true;
        }

        retentionPolicyChanged = !string.Equals(retentionPolicyOverride, _originalSettings.RetentionPolicy, StringComparison.Ordinal);
        return request;
    }

    public void ApplyRetentionPolicyOverride(string policyName)
    {
        _isInitializing = true;
        SelectedRetentionPolicy = policyName;
        _isInitializing = false;
        _retentionPolicyFallback = policyName;
        _originalSettings = CaptureSnapshot();
        HasPendingChanges = false;
    }

    public void DiscardChanges()
    {
        if (_originalSettings == null)
        {
            return;
        }

        _isInitializing = true;
        PrimarySmtpAddress = _originalSettings.PrimarySmtpAddress;
        ProxyAddressesText = string.Join(Environment.NewLine, _originalSettings.ProxyAddresses);
        HiddenFromAddressListsEnabled = _originalSettings.HiddenFromAddressListsEnabled;
        ForwardingAddress = _originalSettings.ForwardingAddress;
        ForwardingSmtpAddress = _originalSettings.ForwardingSmtpAddress;
        DeliverToMailboxAndForward = _originalSettings.DeliverToMailboxAndForward;
        ArchiveEnabled = _originalSettings.ArchiveEnabled;
        LitigationHoldEnabled = _originalSettings.LitigationHoldEnabled;
        AuditEnabled = _originalSettings.AuditEnabled;
        SingleItemRecoveryEnabled = _originalSettings.SingleItemRecoveryEnabled;
        RetentionHoldEnabled = _originalSettings.RetentionHoldEnabled;
        IssueWarningQuota = _originalSettings.IssueWarningQuota;
        ProhibitSendQuota = _originalSettings.ProhibitSendQuota;
        ProhibitSendReceiveQuota = _originalSettings.ProhibitSendReceiveQuota;
        MaxSendSize = _originalSettings.MaxSendSize;
        MaxReceiveSize = _originalSettings.MaxReceiveSize;
        OwaEnabled = _originalSettings.OwaEnabled;
        ActiveSyncEnabled = _originalSettings.ActiveSyncEnabled;
        MapiEnabled = _originalSettings.MapiEnabled;
        PopEnabled = _originalSettings.PopEnabled;
        ImapEnabled = _originalSettings.ImapEnabled;
        SmtpClientAuthenticationDisabled = _originalSettings.SmtpClientAuthenticationDisabled;
        SelectedRetentionPolicy = _originalSettings.RetentionPolicy;
        _isInitializing = false;
        HasPendingChanges = false;
    }

    private void UpdatePendingChanges()
    {
        if (_isInitializing || _isRefreshingRetentionPolicies || _originalSettings == null)
        {
            return;
        }

        HasPendingChanges = !CaptureSnapshot().Equals(_originalSettings);
    }

    private MailboxSettingsSnapshot CaptureSnapshot()
    {
        return new MailboxSettingsSnapshot
        {
            PrimarySmtpAddress = NormalizeInput(PrimarySmtpAddress),
            ProxyAddresses = ParseProxyAddresses(ProxyAddressesText),
            HiddenFromAddressListsEnabled = HiddenFromAddressListsEnabled,
            ForwardingAddress = NormalizeInput(ForwardingAddress),
            ForwardingSmtpAddress = NormalizeInput(ForwardingSmtpAddress),
            DeliverToMailboxAndForward = DeliverToMailboxAndForward,
            ArchiveEnabled = ArchiveEnabled,
            LitigationHoldEnabled = LitigationHoldEnabled,
            AuditEnabled = AuditEnabled,
            SingleItemRecoveryEnabled = SingleItemRecoveryEnabled,
            RetentionHoldEnabled = RetentionHoldEnabled,
            IssueWarningQuota = NormalizeInput(IssueWarningQuota),
            ProhibitSendQuota = NormalizeInput(ProhibitSendQuota),
            ProhibitSendReceiveQuota = NormalizeInput(ProhibitSendReceiveQuota),
            MaxSendSize = NormalizeInput(MaxSendSize),
            MaxReceiveSize = NormalizeInput(MaxReceiveSize),
            RetentionPolicy = NormalizeInput(SelectedRetentionPolicy),
            OwaEnabled = OwaEnabled,
            ActiveSyncEnabled = ActiveSyncEnabled,
            MapiEnabled = MapiEnabled,
            PopEnabled = PopEnabled,
            ImapEnabled = ImapEnabled,
            SmtpClientAuthenticationDisabled = SmtpClientAuthenticationDisabled
        };
    }

    private bool IsPolicySelected(RetentionPolicySummaryDto policy)
    {
        var selected = SelectedRetentionPolicy ?? string.Empty;
        return string.Equals(policy.Name, selected, StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeInput(string? value) => string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();

    private static string BuildProxyAddressesText(IEnumerable<string>? emailAddresses, string? primarySmtpAddress)
    {
        if (emailAddresses == null)
        {
            return string.Empty;
        }

        var primary = NormalizeInput(primarySmtpAddress);
        var additionalAddresses = emailAddresses
            .Where(address => !IsPrimarySmtpAddress(address, primary))
            .Select(NormalizeProxyAddress)
            .Where(address => !string.IsNullOrWhiteSpace(address));

        return string.Join(Environment.NewLine, additionalAddresses);
    }

    private static List<string> ParseProxyAddresses(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return new List<string>();
        }

        return value
            .Split([',', ';', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(NormalizeProxyAddress)
            .Where(address => !string.IsNullOrWhiteSpace(address))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string NormalizeProxyAddress(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var trimmed = value.Trim();
        var separatorIndex = trimmed.IndexOf(':');
        if (separatorIndex <= 0)
        {
            return $"smtp:{trimmed}";
        }

        var prefix = trimmed[..separatorIndex].Trim();
        var address = trimmed[(separatorIndex + 1)..].Trim();
        if (string.IsNullOrWhiteSpace(address))
        {
            return string.Empty;
        }

        return string.Equals(prefix, "smtp", StringComparison.OrdinalIgnoreCase)
            ? $"smtp:{address}"
            : $"{prefix}:{address}";
    }

    private static bool IsPrimarySmtpAddress(string? proxyAddress, string primarySmtpAddress)
    {
        if (string.IsNullOrWhiteSpace(proxyAddress) || string.IsNullOrWhiteSpace(primarySmtpAddress))
        {
            return false;
        }

        var trimmed = proxyAddress.Trim();
        var separatorIndex = trimmed.IndexOf(':');
        if (separatorIndex <= 0)
        {
            return string.Equals(trimmed, primarySmtpAddress, StringComparison.OrdinalIgnoreCase);
        }

        var prefix = trimmed[..separatorIndex].Trim();
        var address = trimmed[(separatorIndex + 1)..].Trim();
        return string.Equals(prefix, "smtp", StringComparison.OrdinalIgnoreCase)
            && string.Equals(address, primarySmtpAddress, StringComparison.OrdinalIgnoreCase);
    }

    private void OnShellViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ShellViewModel.Capabilities))
        {
            ApplyCapabilities(_shellViewModel.Capabilities);
        }
    }
}

