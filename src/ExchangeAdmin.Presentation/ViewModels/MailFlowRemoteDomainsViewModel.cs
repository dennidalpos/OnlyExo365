using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

internal sealed class MailFlowRemoteDomainsViewModel : MailFlowSectionViewModelBase
{
    private static readonly string[] ValidationProperties =
    {
        nameof(IsRemoteDomainInputValid),
        nameof(RemoteDomainValidationMessage)
    };

    private RemoteDomainDto? _selectedRemoteDomain;
    private string? _remoteDomainIdentity;
    private string _remoteDomainName = string.Empty;
    private string _remoteDomainDomainName = string.Empty;
    private string _remoteDomainAllowedOofType = "External";
    private bool _remoteDomainAutoReplyEnabled = true;
    private bool _remoteDomainAutoForwardEnabled = true;
    private bool _remoteDomainDeliveryReportEnabled = true;
    private bool _remoteDomainNdrEnabled = true;
    private bool _remoteDomainMeetingForwardNotificationEnabled = true;
    private bool _remoteDomainTnefEnabled;
    private bool _remoteDomainTrustedMailOutboundEnabled;
    private bool _remoteDomainIsDefault;

    public MailFlowRemoteDomainsViewModel(
        IMailFlowWorkerService workerService,
        ShellViewModel shellViewModel,
        MailFlowOperationCoordinator coordinator,
        Func<CancellationToken, Task> refreshAllAsync)
        : base(workerService, shellViewModel, coordinator, refreshAllAsync)
    {
        NewRemoteDomainCommand = new AsyncRelayCommand(NewRemoteDomainAsync, () => !Coordinator.IsLoading);
        SaveRemoteDomainCommand = new AsyncRelayCommand(SaveRemoteDomainAsync, () => !Coordinator.IsLoading && IsRemoteDomainInputValid);
        RemoveRemoteDomainCommand = new AsyncRelayCommand(RemoveRemoteDomainAsync, () => !Coordinator.IsLoading && SelectedRemoteDomain != null && !SelectedRemoteDomain.IsDefault);
    }

    public IReadOnlyList<string> AllowedOofTypes { get; } = new[] { "External", "ExternalLegacy", "InternalLegacy", "None" };
    public ObservableCollection<RemoteDomainDto> RemoteDomains { get; } = new();

    public RemoteDomainDto? SelectedRemoteDomain
    {
        get => _selectedRemoteDomain;
        set
        {
            if (SetProperty(ref _selectedRemoteDomain, value))
            {
                if (value != null)
                {
                    RemoteDomainIdentity = value.Identity;
                    RemoteDomainName = value.Name;
                    RemoteDomainDomainName = value.DomainName;
                    RemoteDomainAllowedOofType = string.IsNullOrWhiteSpace(value.AllowedOOFType) ? "External" : value.AllowedOOFType;
                    RemoteDomainAutoReplyEnabled = value.AutoReplyEnabled;
                    RemoteDomainAutoForwardEnabled = value.AutoForwardEnabled;
                    RemoteDomainDeliveryReportEnabled = value.DeliveryReportEnabled;
                    RemoteDomainNdrEnabled = value.NDREnabled;
                    RemoteDomainMeetingForwardNotificationEnabled = value.MeetingForwardNotificationEnabled;
                    RemoteDomainTnefEnabled = value.TNEFEnabled;
                    RemoteDomainTrustedMailOutboundEnabled = value.TrustedMailOutboundEnabled;
                    RemoteDomainIsDefault = value.IsDefault;
                }
                else
                {
                    ResetEditor();
                }

                OnPropertyChanged(nameof(CanEditSelectedRemoteDomain));
                InvalidateCommands();
            }
        }
    }

    public bool CanEditSelectedRemoteDomain => SelectedRemoteDomain != null && !Coordinator.IsLoading;
    public bool IsRemoteDomainInputValid =>
        !string.IsNullOrWhiteSpace(RemoteDomainName) &&
        MailFlowViewModelSupport.IsValidRemoteDomain(RemoteDomainDomainName, allowDefaultWildcard: RemoteDomainIdentity != null) &&
        AllowedOofTypes.Contains(RemoteDomainAllowedOofType);

    public string RemoteDomainValidationMessage => IsRemoteDomainInputValid ? string.Empty : "Remote domain: name is required, domain must be valid or wildcard `*.domain.tld`, and AllowedOOFType must be supported.";

    public string? RemoteDomainIdentity { get => _remoteDomainIdentity; set => SetEditorProperty(ref _remoteDomainIdentity, value); }
    public string RemoteDomainName { get => _remoteDomainName; set => SetEditorProperty(ref _remoteDomainName, value); }
    public string RemoteDomainDomainName { get => _remoteDomainDomainName; set => SetEditorProperty(ref _remoteDomainDomainName, value); }
    public string RemoteDomainAllowedOofType { get => _remoteDomainAllowedOofType; set => SetEditorProperty(ref _remoteDomainAllowedOofType, value); }
    public bool RemoteDomainAutoReplyEnabled { get => _remoteDomainAutoReplyEnabled; set => SetProperty(ref _remoteDomainAutoReplyEnabled, value); }
    public bool RemoteDomainAutoForwardEnabled { get => _remoteDomainAutoForwardEnabled; set => SetProperty(ref _remoteDomainAutoForwardEnabled, value); }
    public bool RemoteDomainDeliveryReportEnabled { get => _remoteDomainDeliveryReportEnabled; set => SetProperty(ref _remoteDomainDeliveryReportEnabled, value); }
    public bool RemoteDomainNdrEnabled { get => _remoteDomainNdrEnabled; set => SetProperty(ref _remoteDomainNdrEnabled, value); }
    public bool RemoteDomainMeetingForwardNotificationEnabled { get => _remoteDomainMeetingForwardNotificationEnabled; set => SetProperty(ref _remoteDomainMeetingForwardNotificationEnabled, value); }
    public bool RemoteDomainTnefEnabled { get => _remoteDomainTnefEnabled; set => SetProperty(ref _remoteDomainTnefEnabled, value); }
    public bool RemoteDomainTrustedMailOutboundEnabled { get => _remoteDomainTrustedMailOutboundEnabled; set => SetProperty(ref _remoteDomainTrustedMailOutboundEnabled, value); }
    public bool RemoteDomainIsDefault { get => _remoteDomainIsDefault; set => SetProperty(ref _remoteDomainIsDefault, value); }

    public ICommand NewRemoteDomainCommand { get; }
    public ICommand SaveRemoteDomainCommand { get; }
    public ICommand RemoveRemoteDomainCommand { get; }

    public async Task LoadAsync(CancellationToken cancellationToken)
    {
        var result = await WorkerService.GetRemoteDomainsAsync(new GetRemoteDomainsRequest(), cancellationToken: cancellationToken);
        if (!result.IsSuccess)
        {
            var error = result.Error?.Message ?? "Unable to load remote domains";
            SetError(error);
            ShellViewModel.AddLog(LogLevel.Error, $"MailFlow remote domains load failed: {error}", "MailFlow");
            return;
        }

        RemoteDomains.Clear();
        foreach (var item in result.Value?.Domains ?? new List<RemoteDomainDto>())
        {
            RemoteDomains.Add(item);
        }
    }

    private Task NewRemoteDomainAsync(CancellationToken cancellationToken)
    {
        SelectedRemoteDomain = null;
        ResetEditor();
        SetError(null);
        return Task.CompletedTask;
    }

    private async Task SaveRemoteDomainAsync(CancellationToken cancellationToken)
    {
        if (!IsRemoteDomainInputValid)
        {
            return;
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.UpsertRemoteDomainAsync(new UpsertRemoteDomainRequest
            {
                Identity = string.IsNullOrWhiteSpace(RemoteDomainIdentity) ? null : RemoteDomainIdentity,
                Name = RemoteDomainName.Trim(),
                DomainName = RemoteDomainDomainName.Trim(),
                AllowedOOFType = RemoteDomainAllowedOofType,
                AutoReplyEnabled = RemoteDomainAutoReplyEnabled,
                AutoForwardEnabled = RemoteDomainAutoForwardEnabled,
                DeliveryReportEnabled = RemoteDomainDeliveryReportEnabled,
                NDREnabled = RemoteDomainNdrEnabled,
                MeetingForwardNotificationEnabled = RemoteDomainMeetingForwardNotificationEnabled,
                TNEFEnabled = RemoteDomainTnefEnabled,
                TrustedMailOutboundEnabled = RemoteDomainTrustedMailOutboundEnabled
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                var correlationId = Guid.NewGuid().ToString("N");
                SetError($"Saving remote domain failed (ref: {correlationId}).");
                ShellViewModel.AddLog(LogLevel.Error, $"[{correlationId}] Save remote domain failed (domain={RemoteDomainDomainName}): {result.Error?.Message}", "MailFlow");
                return;
            }

            await RefreshAllAsync(ct);
        }, cancellationToken);
    }

    private async Task RemoveRemoteDomainAsync(CancellationToken cancellationToken)
    {
        if (SelectedRemoteDomain == null || SelectedRemoteDomain.IsDefault)
        {
            return;
        }

        var confirmed = ErrorDialogService.ShowConfirmation("Confirm tenant-wide deletion", $"Operation: Deleting remote domain\nTarget: {SelectedRemoteDomain.DomainName}\nImpact: updates auto-reply and reporting behavior for the remote domain.\n\nConfirm?");
        if (!confirmed)
        {
            return;
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.RemoveRemoteDomainAsync(new RemoveRemoteDomainRequest
            {
                Identity = SelectedRemoteDomain.Identity
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                var error = result.Error?.Message ?? "Unable to delete remote domain";
                SetError(error);
                ShellViewModel.AddLog(LogLevel.Error, $"Remove remote domain failed (domain={SelectedRemoteDomain.DomainName}): {error}", "MailFlow");
                return;
            }

            await RefreshAllAsync(ct);
        }, cancellationToken);
    }

    private void ResetEditor()
    {
        RemoteDomainIdentity = null;
        RemoteDomainName = string.Empty;
        RemoteDomainDomainName = string.Empty;
        RemoteDomainAllowedOofType = "External";
        RemoteDomainAutoReplyEnabled = true;
        RemoteDomainAutoForwardEnabled = true;
        RemoteDomainDeliveryReportEnabled = true;
        RemoteDomainNdrEnabled = true;
        RemoteDomainMeetingForwardNotificationEnabled = true;
        RemoteDomainTnefEnabled = false;
        RemoteDomainTrustedMailOutboundEnabled = false;
        RemoteDomainIsDefault = false;
    }

    private void SetEditorProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (SetProperty(ref field, value, propertyName))
        {
            RaiseProperties(ValidationProperties);
            InvalidateCommands();
        }
    }
}
