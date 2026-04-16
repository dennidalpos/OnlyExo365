using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

internal sealed class MailFlowOrganizationRelationshipsViewModel : MailFlowSectionViewModelBase
{
    private static readonly string[] ValidationProperties =
    {
        nameof(IsOrganizationRelationshipInputValid),
        nameof(OrganizationRelationshipValidationMessage)
    };

    private OrganizationRelationshipDto? _selectedOrganizationRelationship;
    private string? _organizationRelationshipIdentity;
    private string _organizationRelationshipName = string.Empty;
    private string _organizationRelationshipDomainNames = string.Empty;
    private bool _organizationRelationshipEnabled = true;
    private bool _organizationRelationshipFreeBusyAccessEnabled = true;
    private string _organizationRelationshipFreeBusyAccessLevel = "AvailabilityOnly";
    private bool _organizationRelationshipMailTipsAccessEnabled;
    private string _organizationRelationshipMailTipsAccessLevel = "All";
    private string _organizationRelationshipTargetApplicationUri = string.Empty;
    private string _organizationRelationshipTargetAutodiscoverEpr = string.Empty;
    private bool _organizationRelationshipArchiveAccessEnabled;
    private bool _organizationRelationshipDeliveryReportEnabled;
    private bool _organizationRelationshipMailboxMoveEnabled;
    private bool _organizationRelationshipPhotosEnabled;

    public MailFlowOrganizationRelationshipsViewModel(
        IMailFlowWorkerService workerService,
        ShellViewModel shellViewModel,
        MailFlowOperationCoordinator coordinator,
        Func<CancellationToken, Task> refreshAllAsync)
        : base(workerService, shellViewModel, coordinator, refreshAllAsync)
    {
        NewOrganizationRelationshipCommand = new AsyncRelayCommand(NewOrganizationRelationshipAsync, () => !Coordinator.IsLoading);
        SaveOrganizationRelationshipCommand = new AsyncRelayCommand(SaveOrganizationRelationshipAsync, () => !Coordinator.IsLoading && IsOrganizationRelationshipInputValid);
        RemoveOrganizationRelationshipCommand = new AsyncRelayCommand(RemoveOrganizationRelationshipAsync, () => !Coordinator.IsLoading && SelectedOrganizationRelationship != null);
    }

    public IReadOnlyList<string> FreeBusyAccessLevels { get; } = new[] { "AvailabilityOnly", "LimitedDetails", "None" };
    public IReadOnlyList<string> MailTipsAccessLevels { get; } = new[] { "All", "Limited", "None" };
    public ObservableCollection<OrganizationRelationshipDto> OrganizationRelationships { get; } = new();

    public OrganizationRelationshipDto? SelectedOrganizationRelationship
    {
        get => _selectedOrganizationRelationship;
        set
        {
            if (SetProperty(ref _selectedOrganizationRelationship, value))
            {
                if (value != null)
                {
                    OrganizationRelationshipIdentity = value.Identity;
                    OrganizationRelationshipName = value.Name;
                    OrganizationRelationshipDomainNames = string.Join(",", value.DomainNames);
                    OrganizationRelationshipEnabled = value.Enabled;
                    OrganizationRelationshipFreeBusyAccessEnabled = value.FreeBusyAccessEnabled;
                    OrganizationRelationshipFreeBusyAccessLevel = string.IsNullOrWhiteSpace(value.FreeBusyAccessLevel) ? "AvailabilityOnly" : value.FreeBusyAccessLevel;
                    OrganizationRelationshipMailTipsAccessEnabled = value.MailTipsAccessEnabled;
                    OrganizationRelationshipMailTipsAccessLevel = string.IsNullOrWhiteSpace(value.MailTipsAccessLevel) ? "All" : value.MailTipsAccessLevel;
                    OrganizationRelationshipTargetApplicationUri = value.TargetApplicationUri ?? string.Empty;
                    OrganizationRelationshipTargetAutodiscoverEpr = value.TargetAutodiscoverEpr ?? string.Empty;
                    OrganizationRelationshipArchiveAccessEnabled = value.ArchiveAccessEnabled ?? false;
                    OrganizationRelationshipDeliveryReportEnabled = value.DeliveryReportEnabled ?? false;
                    OrganizationRelationshipMailboxMoveEnabled = value.MailboxMoveEnabled ?? false;
                    OrganizationRelationshipPhotosEnabled = value.PhotosEnabled ?? false;
                }
                else
                {
                    ResetEditor();
                }

                OnPropertyChanged(nameof(CanEditSelectedOrganizationRelationship));
                InvalidateCommands();
            }
        }
    }

    public bool CanEditSelectedOrganizationRelationship => SelectedOrganizationRelationship != null && !Coordinator.IsLoading;
    public bool IsOrganizationRelationshipInputValid =>
        !string.IsNullOrWhiteSpace(OrganizationRelationshipName) &&
        MailFlowViewModelSupport.SplitCsv(OrganizationRelationshipDomainNames).Count > 0 &&
        MailFlowViewModelSupport.AreValidDomains(MailFlowViewModelSupport.SplitCsv(OrganizationRelationshipDomainNames)) &&
        FreeBusyAccessLevels.Contains(OrganizationRelationshipFreeBusyAccessLevel) &&
        MailTipsAccessLevels.Contains(OrganizationRelationshipMailTipsAccessLevel) &&
        MailFlowViewModelSupport.IsValidOptionalUri(OrganizationRelationshipTargetApplicationUri) &&
        MailFlowViewModelSupport.IsValidOptionalUri(OrganizationRelationshipTargetAutodiscoverEpr);

    public string OrganizationRelationshipValidationMessage => IsOrganizationRelationshipInputValid ? string.Empty : "Organization relationship: name is required, at least one valid domain is required, supported Free/Busy and MailTips levels must be selected, and optional URIs must use an absolute format.";

    public string? OrganizationRelationshipIdentity { get => _organizationRelationshipIdentity; set => SetEditorProperty(ref _organizationRelationshipIdentity, value); }
    public string OrganizationRelationshipName { get => _organizationRelationshipName; set => SetEditorProperty(ref _organizationRelationshipName, value); }
    public string OrganizationRelationshipDomainNames { get => _organizationRelationshipDomainNames; set => SetEditorProperty(ref _organizationRelationshipDomainNames, value); }
    public bool OrganizationRelationshipEnabled { get => _organizationRelationshipEnabled; set => SetProperty(ref _organizationRelationshipEnabled, value); }
    public bool OrganizationRelationshipFreeBusyAccessEnabled { get => _organizationRelationshipFreeBusyAccessEnabled; set => SetProperty(ref _organizationRelationshipFreeBusyAccessEnabled, value); }
    public string OrganizationRelationshipFreeBusyAccessLevel { get => _organizationRelationshipFreeBusyAccessLevel; set => SetEditorProperty(ref _organizationRelationshipFreeBusyAccessLevel, value); }
    public bool OrganizationRelationshipMailTipsAccessEnabled { get => _organizationRelationshipMailTipsAccessEnabled; set => SetProperty(ref _organizationRelationshipMailTipsAccessEnabled, value); }
    public string OrganizationRelationshipMailTipsAccessLevel { get => _organizationRelationshipMailTipsAccessLevel; set => SetEditorProperty(ref _organizationRelationshipMailTipsAccessLevel, value); }
    public string OrganizationRelationshipTargetApplicationUri { get => _organizationRelationshipTargetApplicationUri; set => SetEditorProperty(ref _organizationRelationshipTargetApplicationUri, value); }
    public string OrganizationRelationshipTargetAutodiscoverEpr { get => _organizationRelationshipTargetAutodiscoverEpr; set => SetEditorProperty(ref _organizationRelationshipTargetAutodiscoverEpr, value); }
    public bool OrganizationRelationshipArchiveAccessEnabled { get => _organizationRelationshipArchiveAccessEnabled; set => SetProperty(ref _organizationRelationshipArchiveAccessEnabled, value); }
    public bool OrganizationRelationshipDeliveryReportEnabled { get => _organizationRelationshipDeliveryReportEnabled; set => SetProperty(ref _organizationRelationshipDeliveryReportEnabled, value); }
    public bool OrganizationRelationshipMailboxMoveEnabled { get => _organizationRelationshipMailboxMoveEnabled; set => SetProperty(ref _organizationRelationshipMailboxMoveEnabled, value); }
    public bool OrganizationRelationshipPhotosEnabled { get => _organizationRelationshipPhotosEnabled; set => SetProperty(ref _organizationRelationshipPhotosEnabled, value); }

    public ICommand NewOrganizationRelationshipCommand { get; }
    public ICommand SaveOrganizationRelationshipCommand { get; }
    public ICommand RemoveOrganizationRelationshipCommand { get; }

    public async Task LoadAsync(CancellationToken cancellationToken)
    {
        var result = await WorkerService.GetOrganizationRelationshipsAsync(new GetOrganizationRelationshipsRequest(), cancellationToken: cancellationToken);
        if (!result.IsSuccess)
        {
            var error = result.Error?.Message ?? "Unable to load organization relationships";
            SetError(error);
            ShellViewModel.AddLog(LogLevel.Error, $"MailFlow organization relationships load failed: {error}", "MailFlow");
            return;
        }

        OrganizationRelationships.Clear();
        foreach (var item in result.Value?.Relationships ?? new List<OrganizationRelationshipDto>())
        {
            OrganizationRelationships.Add(item);
        }
    }

    private Task NewOrganizationRelationshipAsync(CancellationToken cancellationToken)
    {
        SelectedOrganizationRelationship = null;
        ResetEditor();
        SetError(null);
        return Task.CompletedTask;
    }

    private async Task SaveOrganizationRelationshipAsync(CancellationToken cancellationToken)
    {
        if (!IsOrganizationRelationshipInputValid)
        {
            return;
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.UpsertOrganizationRelationshipAsync(new UpsertOrganizationRelationshipRequest
            {
                Identity = string.IsNullOrWhiteSpace(OrganizationRelationshipIdentity) ? null : OrganizationRelationshipIdentity,
                Name = OrganizationRelationshipName.Trim(),
                DomainNames = MailFlowViewModelSupport.SplitCsv(OrganizationRelationshipDomainNames),
                Enabled = OrganizationRelationshipEnabled,
                FreeBusyAccessEnabled = OrganizationRelationshipFreeBusyAccessEnabled,
                FreeBusyAccessLevel = OrganizationRelationshipFreeBusyAccessLevel,
                MailTipsAccessEnabled = OrganizationRelationshipMailTipsAccessEnabled,
                MailTipsAccessLevel = OrganizationRelationshipMailTipsAccessLevel,
                TargetApplicationUri = string.IsNullOrWhiteSpace(OrganizationRelationshipTargetApplicationUri) ? null : OrganizationRelationshipTargetApplicationUri.Trim(),
                TargetAutodiscoverEpr = string.IsNullOrWhiteSpace(OrganizationRelationshipTargetAutodiscoverEpr) ? null : OrganizationRelationshipTargetAutodiscoverEpr.Trim(),
                ArchiveAccessEnabled = OrganizationRelationshipArchiveAccessEnabled,
                DeliveryReportEnabled = OrganizationRelationshipDeliveryReportEnabled,
                MailboxMoveEnabled = OrganizationRelationshipMailboxMoveEnabled,
                PhotosEnabled = OrganizationRelationshipPhotosEnabled
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                var correlationId = Guid.NewGuid().ToString("N");
                SetError($"Saving organization relationship failed (ref: {correlationId}).");
                ShellViewModel.AddLog(LogLevel.Error, $"[{correlationId}] Save organization relationship failed (name={OrganizationRelationshipName}): {result.Error?.Message}", "MailFlow");
                return;
            }

            await RefreshAllAsync(ct);
        }, cancellationToken);
    }

    private async Task RemoveOrganizationRelationshipAsync(CancellationToken cancellationToken)
    {
        if (SelectedOrganizationRelationship == null)
        {
            return;
        }

        var confirmed = ErrorDialogService.ShowConfirmation("Confirm tenant-wide deletion", $"Operation: Deleting organization relationship\nTarget: {SelectedOrganizationRelationship.Name}\nImpact: can interrupt free/busy, MailTips, or cross-tenant integrations.\n\nConfirm?");
        if (!confirmed)
        {
            return;
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.RemoveOrganizationRelationshipAsync(new RemoveOrganizationRelationshipRequest
            {
                Identity = SelectedOrganizationRelationship.Identity
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                var error = result.Error?.Message ?? "Unable to delete organization relationship";
                SetError(error);
                ShellViewModel.AddLog(LogLevel.Error, $"Remove organization relationship failed (name={SelectedOrganizationRelationship.Name}): {error}", "MailFlow");
                return;
            }

            await RefreshAllAsync(ct);
        }, cancellationToken);
    }

    private void ResetEditor()
    {
        OrganizationRelationshipIdentity = null;
        OrganizationRelationshipName = string.Empty;
        OrganizationRelationshipDomainNames = string.Empty;
        OrganizationRelationshipEnabled = true;
        OrganizationRelationshipFreeBusyAccessEnabled = true;
        OrganizationRelationshipFreeBusyAccessLevel = "AvailabilityOnly";
        OrganizationRelationshipMailTipsAccessEnabled = false;
        OrganizationRelationshipMailTipsAccessLevel = "All";
        OrganizationRelationshipTargetApplicationUri = string.Empty;
        OrganizationRelationshipTargetAutodiscoverEpr = string.Empty;
        OrganizationRelationshipArchiveAccessEnabled = false;
        OrganizationRelationshipDeliveryReportEnabled = false;
        OrganizationRelationshipMailboxMoveEnabled = false;
        OrganizationRelationshipPhotosEnabled = false;
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
