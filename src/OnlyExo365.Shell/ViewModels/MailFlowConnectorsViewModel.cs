using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Shell.Helpers;

namespace OnlyExo365.Shell.ViewModels;

internal sealed class MailFlowConnectorsViewModel : MailFlowSectionViewModelBase
{
    private static readonly string[] ValidationProperties =
    {
        nameof(IsConnectorInputValid),
        nameof(ConnectorValidationMessage)
    };

    private ConnectorDto? _selectedConnector;
    private string? _connectorIdentity;
    private string? _connectorIdentityDisplay;
    private string _connectorType = "Inbound";
    private string _connectorName = string.Empty;
    private string _connectorComment = string.Empty;
    private bool _connectorEnabled = true;
    private string _connectorSenderDomains = string.Empty;
    private string _connectorRecipientDomains = string.Empty;

    public MailFlowConnectorsViewModel(
        IMailFlowWorkerService workerService,
        ShellViewModel shellViewModel,
        MailFlowOperationCoordinator coordinator,
        Func<CancellationToken, Task> refreshAllAsync)
        : base(workerService, shellViewModel, coordinator, refreshAllAsync)
    {
        NewConnectorCommand = new RelayCommand(BeginCreateConnector, () => !Coordinator.IsLoading);
        SaveConnectorCommand = new AsyncRelayCommand(SaveConnectorAsync, () => !Coordinator.IsLoading && IsConnectorInputValid);
        RemoveConnectorCommand = new AsyncRelayCommand(RemoveConnectorAsync, () => !Coordinator.IsLoading && SelectedConnector != null);
    }

    public IReadOnlyList<string> ConnectorTypes { get; } = new[] { "Inbound", "Outbound" };
    public ObservableCollection<ConnectorDto> Connectors { get; } = new();

    public ConnectorDto? SelectedConnector
    {
        get => _selectedConnector;
        set
        {
            if (SetProperty(ref _selectedConnector, value))
            {
                if (value != null)
                {
                    ConnectorIdentity = value.Identity;
                    ConnectorIdentityDisplay = string.IsNullOrWhiteSpace(value.DisplayLabel) ? value.Name : value.DisplayLabel;
                    ConnectorType = string.IsNullOrWhiteSpace(value.Type) ? "Inbound" : value.Type;
                    ConnectorName = value.Name;
                    ConnectorComment = value.Comment;
                    ConnectorEnabled = value.Enabled;
                    ConnectorSenderDomains = string.Join(",", value.SenderDomains);
                    ConnectorRecipientDomains = string.Join(",", value.RecipientDomains);
                }
                else
                {
                    ResetEditor();
                }

                OnPropertyChanged(nameof(CanEditSelectedConnector));
                InvalidateCommands();
            }
        }
    }

    public bool CanEditSelectedConnector => SelectedConnector != null && !Coordinator.IsLoading;
    public bool IsConnectorInputValid =>
        !string.IsNullOrWhiteSpace(ConnectorName) &&
        ConnectorTypes.Contains(ConnectorType) &&
        MailFlowViewModelSupport.AreValidDomains(MailFlowViewModelSupport.SplitCsv(ConnectorSenderDomains)) &&
        MailFlowViewModelSupport.AreValidDomains(MailFlowViewModelSupport.SplitCsv(ConnectorRecipientDomains));

    public string ConnectorValidationMessage => IsConnectorInputValid ? string.Empty : "Connector: name is required, type must be valid, and domains must use the correct format.";

    public string? ConnectorIdentity { get => _connectorIdentity; set => SetProperty(ref _connectorIdentity, value); }
    public string? ConnectorIdentityDisplay { get => _connectorIdentityDisplay; set => SetProperty(ref _connectorIdentityDisplay, value); }
    public string ConnectorType { get => _connectorType; set => SetEditorProperty(ref _connectorType, value); }
    public string ConnectorName { get => _connectorName; set => SetEditorProperty(ref _connectorName, value); }
    public string ConnectorComment { get => _connectorComment; set => SetProperty(ref _connectorComment, value); }
    public bool ConnectorEnabled { get => _connectorEnabled; set => SetProperty(ref _connectorEnabled, value); }
    public string ConnectorSenderDomains { get => _connectorSenderDomains; set => SetEditorProperty(ref _connectorSenderDomains, value); }
    public string ConnectorRecipientDomains { get => _connectorRecipientDomains; set => SetEditorProperty(ref _connectorRecipientDomains, value); }

    public ICommand NewConnectorCommand { get; }
    public ICommand SaveConnectorCommand { get; }
    public ICommand RemoveConnectorCommand { get; }

    public async Task LoadAsync(CancellationToken cancellationToken)
    {
        var result = await WorkerService.GetConnectorsAsync(new GetConnectorsRequest(), cancellationToken: cancellationToken);
        if (!result.IsSuccess)
        {
            var error = result.Error?.Message ?? "Unable to load connector";
            SetError(error);
            ShellViewModel.AddLog(LogLevel.Error, $"MailFlow connectors load failed: {error}", "MailFlow");
            return;
        }

        Connectors.Clear();
        foreach (var item in result.Value?.Connectors ?? new List<ConnectorDto>())
        {
            Connectors.Add(item);
        }
    }

    private async Task SaveConnectorAsync(CancellationToken cancellationToken)
    {
        if (!IsConnectorInputValid)
        {
            return;
        }

        if (SelectedConnector != null && SelectedConnector.Enabled && !ConnectorEnabled)
        {
            var disableConfirmed = ErrorDialogService.ShowConfirmation(
                "Confirm tenant-wide disable",
                $"Operation: DisEnablezione connector {SelectedConnector.Type}\n" +
                $"Target: {SelectedConnector.Name}\n" +
                "Impact: can interrupt tenant-wide mail flow.\n\n" +
                "Confirm?");

            if (!disableConfirmed)
            {
                return;
            }
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.UpsertConnectorAsync(new UpsertConnectorRequest
            {
                Identity = string.IsNullOrWhiteSpace(ConnectorIdentity) ? null : ConnectorIdentity,
                Type = ConnectorType,
                Name = ConnectorName.Trim(),
                Comment = ConnectorComment,
                Enabled = ConnectorEnabled,
                SenderDomains = MailFlowViewModelSupport.SplitCsv(ConnectorSenderDomains),
                RecipientDomains = MailFlowViewModelSupport.SplitCsv(ConnectorRecipientDomains)
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                var correlationId = Guid.NewGuid().ToString("N");
                SetError($"Saving connector failed. Check the fields and try again (ref: {correlationId}).");
                ShellViewModel.AddLog(LogLevel.Error, $"[{correlationId}] Save connector failed (name={ConnectorName}, type={ConnectorType}): {result.Error?.Message}", "MailFlow");
                return;
            }

            await RefreshAllAsync(ct);
        }, cancellationToken);
    }

    private async Task RemoveConnectorAsync(CancellationToken cancellationToken)
    {
        if (SelectedConnector == null)
        {
            return;
        }

        var confirmed = ErrorDialogService.ShowConfirmation("Confirm tenant-wide deletion", $"Operation: Deleting connector {SelectedConnector.Type}\nTarget: {SelectedConnector.Name}\nImpact: update routing posta tenant-wide.\n\nConfirm?");
        if (!confirmed)
        {
            return;
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.RemoveConnectorAsync(new RemoveConnectorRequest
            {
                Identity = SelectedConnector.Identity,
                Type = SelectedConnector.Type
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                var error = result.Error?.Message ?? "Unable to delete connector";
                SetError(error);
                ShellViewModel.AddLog(LogLevel.Error, $"Remove connector failed (connector={SelectedConnector.Name}): {error}", "MailFlow");
                return;
            }

            await RefreshAllAsync(ct);
        }, cancellationToken);
    }

    private void BeginCreateConnector()
    {
        SelectedConnector = null;
        ResetEditor();
        InvalidateCommands();
    }

    private void ResetEditor()
    {
        ConnectorIdentity = null;
        ConnectorIdentityDisplay = null;
        ConnectorType = "Inbound";
        ConnectorName = string.Empty;
        ConnectorComment = string.Empty;
        ConnectorEnabled = true;
        ConnectorSenderDomains = string.Empty;
        ConnectorRecipientDomains = string.Empty;
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

