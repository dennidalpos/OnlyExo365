using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

internal sealed class MailFlowAcceptedDomainsViewModel : MailFlowSectionViewModelBase
{
    private static readonly string[] ValidationProperties =
    {
        nameof(IsDomainInputValid),
        nameof(DomainValidationMessage)
    };

    private AcceptedDomainDto? _selectedDomain;
    private string? _domainIdentity;
    private string _domainName = string.Empty;
    private string _domainFqdn = string.Empty;
    private string _domainType = "Authoritative";
    private bool _domainMakeDefault;

    public MailFlowAcceptedDomainsViewModel(
        IMailFlowWorkerService workerService,
        ShellViewModel shellViewModel,
        MailFlowOperationCoordinator coordinator,
        Func<CancellationToken, Task> refreshAllAsync)
        : base(workerService, shellViewModel, coordinator, refreshAllAsync)
    {
        NewDomainCommand = new RelayCommand(BeginCreateDomain, () => !Coordinator.IsLoading);
        SaveDomainCommand = new AsyncRelayCommand(SaveDomainAsync, () => !Coordinator.IsLoading && IsDomainInputValid);
        RemoveDomainCommand = new AsyncRelayCommand(RemoveDomainAsync, () => !Coordinator.IsLoading && SelectedDomain != null);
    }

    public IReadOnlyList<string> DomainTypes { get; } = new[] { "Authoritative", "InternalRelay", "ExternalRelay" };
    public ObservableCollection<AcceptedDomainDto> AcceptedDomains { get; } = new();

    public AcceptedDomainDto? SelectedDomain
    {
        get => _selectedDomain;
        set
        {
            if (SetProperty(ref _selectedDomain, value))
            {
                if (value != null)
                {
                    DomainIdentity = value.Identity;
                    DomainName = value.Name;
                    DomainFqdn = value.DomainName;
                    DomainType = value.DomainType;
                    DomainMakeDefault = value.Default;
                }
                else
                {
                    ResetEditor();
                }

                OnPropertyChanged(nameof(CanEditSelectedDomain));
                InvalidateCommands();
            }
        }
    }

    public bool CanEditSelectedDomain => SelectedDomain != null && !Coordinator.IsLoading;
    public bool IsDomainInputValid =>
        !string.IsNullOrWhiteSpace(DomainName) &&
        MailFlowViewModelSupport.IsValidDomain(DomainFqdn) &&
        DomainTypes.Contains(DomainType);

    public string DomainValidationMessage => IsDomainInputValid ? string.Empty : "Domain: name is required, FQDN must be valid, and DomainType must be supported.";

    public string? DomainIdentity { get => _domainIdentity; set => SetProperty(ref _domainIdentity, value); }
    public string DomainName { get => _domainName; set => SetEditorProperty(ref _domainName, value); }
    public string DomainFqdn { get => _domainFqdn; set => SetEditorProperty(ref _domainFqdn, value); }
    public string DomainType { get => _domainType; set => SetEditorProperty(ref _domainType, value); }
    public bool DomainMakeDefault { get => _domainMakeDefault; set => SetProperty(ref _domainMakeDefault, value); }

    public ICommand NewDomainCommand { get; }
    public ICommand SaveDomainCommand { get; }
    public ICommand RemoveDomainCommand { get; }

    public async Task LoadAsync(CancellationToken cancellationToken)
    {
        var result = await WorkerService.GetAcceptedDomainsAsync(new GetAcceptedDomainsRequest(), cancellationToken: cancellationToken);
        if (!result.IsSuccess)
        {
            var error = result.Error?.Message ?? "Unable to load accepted domains";
            SetError(error);
            ShellViewModel.AddLog(LogLevel.Error, $"MailFlow accepted domains load failed: {error}", "MailFlow");
            return;
        }

        AcceptedDomains.Clear();
        foreach (var item in result.Value?.Domains ?? new List<AcceptedDomainDto>())
        {
            AcceptedDomains.Add(item);
        }
    }

    private async Task SaveDomainAsync(CancellationToken cancellationToken)
    {
        if (!IsDomainInputValid)
        {
            return;
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.UpsertAcceptedDomainAsync(new UpsertAcceptedDomainRequest
            {
                Identity = string.IsNullOrWhiteSpace(DomainIdentity) ? null : DomainIdentity,
                Name = DomainName.Trim(),
                DomainName = DomainFqdn.Trim(),
                DomainType = DomainType,
                MakeDefault = DomainMakeDefault
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                var correlationId = Guid.NewGuid().ToString("N");
                SetError($"Saving domain failed (ref: {correlationId}).");
                ShellViewModel.AddLog(LogLevel.Error, $"[{correlationId}] Save domain failed (domain={DomainFqdn}): {result.Error?.Message}", "MailFlow");
                return;
            }

            await RefreshAllAsync(ct);
        }, cancellationToken);
    }

    private async Task RemoveDomainAsync(CancellationToken cancellationToken)
    {
        if (SelectedDomain == null)
        {
            return;
        }

        var confirmed = ErrorDialogService.ShowConfirmation("Confirm tenant-wide deletion", $"Operation: Deleting accepted domain\nTarget: {SelectedDomain.DomainName}\nImpact: can interrupt tenant-wide mail delivery/routing.\n\nConfirm?");
        if (!confirmed)
        {
            return;
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.RemoveAcceptedDomainAsync(new RemoveAcceptedDomainRequest
            {
                Identity = SelectedDomain.Identity
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                var error = result.Error?.Message ?? "Unable to delete domain";
                SetError(error);
                ShellViewModel.AddLog(LogLevel.Error, $"Remove domain failed (domain={SelectedDomain.DomainName}): {error}", "MailFlow");
                return;
            }

            await RefreshAllAsync(ct);
        }, cancellationToken);
    }

    private void BeginCreateDomain()
    {
        SelectedDomain = null;
        ResetEditor();
        InvalidateCommands();
    }

    private void ResetEditor()
    {
        DomainIdentity = null;
        DomainName = string.Empty;
        DomainFqdn = string.Empty;
        DomainType = "Authoritative";
        DomainMakeDefault = false;
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
