using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Shell.Helpers;

namespace OnlyExo365.Shell.ViewModels;

internal sealed class MailFlowSharingPoliciesViewModel : MailFlowSectionViewModelBase
{
    private static readonly string[] ValidationProperties =
    {
        nameof(IsSharingPolicyInputValid),
        nameof(SharingPolicyValidationMessage)
    };

    private SharingPolicyDto? _selectedSharingPolicy;
    private string? _sharingPolicyIdentity;
    private string _sharingPolicyName = string.Empty;
    private string _sharingPolicyDomains = string.Empty;
    private bool _sharingPolicyEnabled = true;
    private bool _sharingPolicyMakeDefault;
    private bool _sharingPolicyIsDefault;

    public MailFlowSharingPoliciesViewModel(
        IMailFlowWorkerService workerService,
        ShellViewModel shellViewModel,
        MailFlowOperationCoordinator coordinator,
        Func<CancellationToken, Task> refreshAllAsync)
        : base(workerService, shellViewModel, coordinator, refreshAllAsync)
    {
        NewSharingPolicyCommand = new AsyncRelayCommand(NewSharingPolicyAsync, () => !Coordinator.IsLoading);
        SaveSharingPolicyCommand = new AsyncRelayCommand(SaveSharingPolicyAsync, () => !Coordinator.IsLoading && IsSharingPolicyInputValid);
        RemoveSharingPolicyCommand = new AsyncRelayCommand(RemoveSharingPolicyAsync, () => !Coordinator.IsLoading && SelectedSharingPolicy != null && !SelectedSharingPolicy.IsDefault);
    }

    public ObservableCollection<SharingPolicyDto> SharingPolicies { get; } = new();

    public SharingPolicyDto? SelectedSharingPolicy
    {
        get => _selectedSharingPolicy;
        set
        {
            if (SetProperty(ref _selectedSharingPolicy, value))
            {
                if (value != null)
                {
                    SharingPolicyIdentity = value.Identity;
                    SharingPolicyName = value.Name;
                    SharingPolicyDomains = string.Join(",", value.Domains);
                    SharingPolicyEnabled = value.Enabled;
                    SharingPolicyIsDefault = value.IsDefault;
                    SharingPolicyMakeDefault = value.IsDefault;
                }
                else
                {
                    ResetEditor();
                }

                OnPropertyChanged(nameof(CanEditSelectedSharingPolicy));
                InvalidateCommands();
            }
        }
    }

    public bool CanEditSelectedSharingPolicy => SelectedSharingPolicy != null && !Coordinator.IsLoading;
    public bool IsSharingPolicyInputValid =>
        !string.IsNullOrWhiteSpace(SharingPolicyName) &&
        MailFlowViewModelSupport.SplitCsv(SharingPolicyDomains).Count > 0;

    public string SharingPolicyValidationMessage
        => IsSharingPolicyInputValid
            ? string.Empty
            : "Sharing policy: Name is required and at least one Domains entry is required (for example, domain.tld: CalendarSharingFreeBusyDetail).";

    public string? SharingPolicyIdentity { get => _sharingPolicyIdentity; set => SetEditorProperty(ref _sharingPolicyIdentity, value); }
    public string SharingPolicyName { get => _sharingPolicyName; set => SetEditorProperty(ref _sharingPolicyName, value); }
    public string SharingPolicyDomains { get => _sharingPolicyDomains; set => SetEditorProperty(ref _sharingPolicyDomains, value); }
    public bool SharingPolicyEnabled { get => _sharingPolicyEnabled; set => SetProperty(ref _sharingPolicyEnabled, value); }
    public bool SharingPolicyMakeDefault { get => _sharingPolicyMakeDefault; set => SetProperty(ref _sharingPolicyMakeDefault, value); }
    public bool SharingPolicyIsDefault { get => _sharingPolicyIsDefault; set => SetProperty(ref _sharingPolicyIsDefault, value); }

    public ICommand NewSharingPolicyCommand { get; }
    public ICommand SaveSharingPolicyCommand { get; }
    public ICommand RemoveSharingPolicyCommand { get; }

    public async Task LoadAsync(CancellationToken cancellationToken)
    {
        var result = await WorkerService.GetSharingPoliciesAsync(new GetSharingPoliciesRequest(), cancellationToken: cancellationToken);
        if (!result.IsSuccess)
        {
            var error = result.Error?.Message ?? "Unable to load sharing policies";
            SetError(error);
            ShellViewModel.AddLog(LogLevel.Error, $"MailFlow sharing policies load failed: {error}", "MailFlow");
            return;
        }

        SharingPolicies.Clear();
        foreach (var item in result.Value?.Policies ?? new List<SharingPolicyDto>())
        {
            SharingPolicies.Add(item);
        }
    }

    private Task NewSharingPolicyAsync(CancellationToken cancellationToken)
    {
        SelectedSharingPolicy = null;
        ResetEditor();
        SetError(null);
        return Task.CompletedTask;
    }

    private async Task SaveSharingPolicyAsync(CancellationToken cancellationToken)
    {
        if (!IsSharingPolicyInputValid)
        {
            return;
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.UpsertSharingPolicyAsync(new UpsertSharingPolicyRequest
            {
                Identity = string.IsNullOrWhiteSpace(SharingPolicyIdentity) ? null : SharingPolicyIdentity,
                Name = SharingPolicyName.Trim(),
                Domains = MailFlowViewModelSupport.SplitCsv(SharingPolicyDomains),
                Enabled = SharingPolicyEnabled,
                MakeDefault = SharingPolicyMakeDefault
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                var correlationId = Guid.NewGuid().ToString("N");
                SetError($"Saving sharing policy failed (ref: {correlationId}).");
                ShellViewModel.AddLog(LogLevel.Error, $"[{correlationId}] Save sharing policy failed (name={SharingPolicyName}): {result.Error?.Message}", "MailFlow");
                return;
            }

            await RefreshAllAsync(ct);
        }, cancellationToken);
    }

    private async Task RemoveSharingPolicyAsync(CancellationToken cancellationToken)
    {
        if (SelectedSharingPolicy == null || SelectedSharingPolicy.IsDefault)
        {
            return;
        }

        var confirmed = ErrorDialogService.ShowConfirmation("Confirm tenant-wide deletion", $"Operation: Deleting sharing policy\nTarget: {SelectedSharingPolicy.Name}\nImpact: can update calendar or contacts sharing rules toward external domains.\n\nConfirm?");
        if (!confirmed)
        {
            return;
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.RemoveSharingPolicyAsync(new RemoveSharingPolicyRequest
            {
                Identity = SelectedSharingPolicy.Identity
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                var error = result.Error?.Message ?? "Unable to delete sharing policy";
                SetError(error);
                ShellViewModel.AddLog(LogLevel.Error, $"Remove sharing policy failed (name={SelectedSharingPolicy.Name}): {error}", "MailFlow");
                return;
            }

            await RefreshAllAsync(ct);
        }, cancellationToken);
    }

    private void ResetEditor()
    {
        SharingPolicyIdentity = null;
        SharingPolicyName = string.Empty;
        SharingPolicyDomains = string.Empty;
        SharingPolicyEnabled = true;
        SharingPolicyMakeDefault = false;
        SharingPolicyIsDefault = false;
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

