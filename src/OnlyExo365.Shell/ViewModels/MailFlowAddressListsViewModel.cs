using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Shell.Helpers;

namespace OnlyExo365.Shell.ViewModels;

internal sealed class MailFlowAddressListsViewModel : MailFlowSectionViewModelBase
{
    private static readonly string[] ValidationProperties =
    {
        nameof(IsAddressListInputValid),
        nameof(AddressListValidationMessage)
    };

    private static readonly HashSet<string> AllowedIncludedRecipients = new(StringComparer.OrdinalIgnoreCase)
    {
        "AllRecipients",
        "MailboxUsers",
        "MailContacts",
        "MailGroups",
        "MailUsers",
        "Resources",
        "RoomMailboxes",
        "EquipmentMailboxes"
    };

    private AddressListDto? _selectedAddressList;
    private string? _addressListIdentity;
    private string _addressListName = string.Empty;
    private string _addressListDisplayName = string.Empty;
    private string _addressListRecipientFilter = string.Empty;
    private string _addressListRecipientContainer = string.Empty;
    private string _addressListIncludedRecipients = string.Empty;
    private string _addressListConditionalCompany = string.Empty;
    private string _addressListConditionalDepartment = string.Empty;
    private string _addressListConditionalStateOrProvince = string.Empty;
    private string _addressListConditionalCustomAttribute1 = string.Empty;

    public MailFlowAddressListsViewModel(
        IMailFlowWorkerService workerService,
        ShellViewModel shellViewModel,
        MailFlowOperationCoordinator coordinator,
        Func<CancellationToken, Task> refreshAllAsync)
        : base(workerService, shellViewModel, coordinator, refreshAllAsync)
    {
        NewAddressListCommand = new AsyncRelayCommand(NewAddressListAsync, () => IsSectionSupported && !Coordinator.IsLoading);
        SaveAddressListCommand = new AsyncRelayCommand(SaveAddressListAsync, () => IsSectionSupported && !Coordinator.IsLoading && IsAddressListInputValid);
        RemoveAddressListCommand = new AsyncRelayCommand(RemoveAddressListAsync, () => IsSectionSupported && !Coordinator.IsLoading && SelectedAddressList != null);
    }

    public ObservableCollection<AddressListDto> AddressLists { get; } = new();

    public AddressListDto? SelectedAddressList
    {
        get => _selectedAddressList;
        set
        {
            if (SetProperty(ref _selectedAddressList, value))
            {
                if (value != null)
                {
                    AddressListIdentity = value.Identity;
                    AddressListName = value.Name;
                    AddressListDisplayName = value.DisplayName;
                    AddressListRecipientFilter = value.RecipientFilter;
                    AddressListRecipientContainer = value.RecipientContainer ?? string.Empty;
                    AddressListIncludedRecipients = string.Join(",", value.IncludedRecipients);
                    AddressListConditionalCompany = string.Join(",", value.ConditionalCompany);
                    AddressListConditionalDepartment = string.Join(",", value.ConditionalDepartment);
                    AddressListConditionalStateOrProvince = string.Join(",", value.ConditionalStateOrProvince);
                    AddressListConditionalCustomAttribute1 = string.Join(",", value.ConditionalCustomAttribute1);
                }
                else
                {
                    ResetEditor();
                }

                OnPropertyChanged(nameof(CanEditSelectedAddressList));
                InvalidateCommands();
            }
        }
    }

    public bool CanEditSelectedAddressList => SelectedAddressList != null && IsSectionSupported && !Coordinator.IsLoading;

    public bool IsAddressListInputValid
    {
        get
        {
            if (string.IsNullOrWhiteSpace(AddressListName))
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(AddressListRecipientFilter))
            {
                return true;
            }

            var includedRecipients = MailFlowViewModelSupport.SplitCsv(AddressListIncludedRecipients);
            return includedRecipients.Count > 0 &&
                   includedRecipients.All(value => AllowedIncludedRecipients.Contains(value));
        }
    }

    public string AddressListValidationMessage
        => IsAddressListInputValid
            ? string.Empty
            : "Address list: Name is required and either a custom RecipientFilter or valid IncludedRecipients (AllRecipients, MailboxUsers, MailContacts, MailGroups, MailUsers, Resources, RoomMailboxes, EquipmentMailboxes) must be provided.";

    public string? AddressListIdentity { get => _addressListIdentity; set => SetEditorProperty(ref _addressListIdentity, value); }
    public string AddressListName { get => _addressListName; set => SetEditorProperty(ref _addressListName, value); }
    public string AddressListDisplayName { get => _addressListDisplayName; set => SetEditorProperty(ref _addressListDisplayName, value); }
    public string AddressListRecipientFilter { get => _addressListRecipientFilter; set => SetEditorProperty(ref _addressListRecipientFilter, value); }
    public string AddressListRecipientContainer { get => _addressListRecipientContainer; set => SetEditorProperty(ref _addressListRecipientContainer, value); }
    public string AddressListIncludedRecipients { get => _addressListIncludedRecipients; set => SetEditorProperty(ref _addressListIncludedRecipients, value); }
    public string AddressListConditionalCompany { get => _addressListConditionalCompany; set => SetEditorProperty(ref _addressListConditionalCompany, value); }
    public string AddressListConditionalDepartment { get => _addressListConditionalDepartment; set => SetEditorProperty(ref _addressListConditionalDepartment, value); }
    public string AddressListConditionalStateOrProvince { get => _addressListConditionalStateOrProvince; set => SetEditorProperty(ref _addressListConditionalStateOrProvince, value); }
    public string AddressListConditionalCustomAttribute1 { get => _addressListConditionalCustomAttribute1; set => SetEditorProperty(ref _addressListConditionalCustomAttribute1, value); }

    public ICommand NewAddressListCommand { get; }
    public ICommand SaveAddressListCommand { get; }
    public ICommand RemoveAddressListCommand { get; }

    public async Task LoadAsync(CancellationToken cancellationToken)
    {
        var result = await WorkerService.GetAddressListsAsync(new GetAddressListsRequest(), cancellationToken: cancellationToken);
        if (!result.IsSuccess)
        {
            var error = result.Error?.Message ?? "Unable to load address lists";
            SetError(error);
            ShellViewModel.AddLog(LogLevel.Error, $"MailFlow address lists load failed: {error}", "MailFlow");
            return;
        }

        AddressLists.Clear();
        var warningMessage = result.Value?.Warnings.Count > 0
            ? string.Join(Environment.NewLine, result.Value.Warnings)
            : null;
        SetSectionState(!(result.Value?.IsUnsupported ?? false), warningMessage);

        if (result.Value?.Warnings != null)
        {
            foreach (var warning in result.Value.Warnings)
            {
                ShellViewModel.AddLog(LogLevel.Warning, warning, "MailFlow", result.CorrelationId);
            }
        }

        if (result.Value?.IsUnsupported ?? false)
        {
            SelectedAddressList = null;
            return;
        }

        foreach (var item in result.Value?.AddressLists ?? new List<AddressListDto>())
        {
            AddressLists.Add(item);
        }
    }

    private Task NewAddressListAsync(CancellationToken cancellationToken)
    {
        SelectedAddressList = null;
        ResetEditor();
        SetError(null);
        return Task.CompletedTask;
    }

    private async Task SaveAddressListAsync(CancellationToken cancellationToken)
    {
        if (!IsAddressListInputValid)
        {
            return;
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.UpsertAddressListAsync(new UpsertAddressListRequest
            {
                Identity = string.IsNullOrWhiteSpace(AddressListIdentity) ? null : AddressListIdentity,
                Name = AddressListName.Trim(),
                DisplayName = string.IsNullOrWhiteSpace(AddressListDisplayName) ? null : AddressListDisplayName.Trim(),
                RecipientFilter = string.IsNullOrWhiteSpace(AddressListRecipientFilter) ? null : AddressListRecipientFilter.Trim(),
                RecipientContainer = string.IsNullOrWhiteSpace(AddressListRecipientContainer) ? null : AddressListRecipientContainer.Trim(),
                IncludedRecipients = MailFlowViewModelSupport.SplitCsv(AddressListIncludedRecipients),
                ConditionalCompany = MailFlowViewModelSupport.SplitCsv(AddressListConditionalCompany),
                ConditionalDepartment = MailFlowViewModelSupport.SplitCsv(AddressListConditionalDepartment),
                ConditionalStateOrProvince = MailFlowViewModelSupport.SplitCsv(AddressListConditionalStateOrProvince),
                ConditionalCustomAttribute1 = MailFlowViewModelSupport.SplitCsv(AddressListConditionalCustomAttribute1)
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                var correlationId = Guid.NewGuid().ToString("N");
                SetError($"Saving address list failed (ref: {correlationId}).");
                ShellViewModel.AddLog(LogLevel.Error, $"[{correlationId}] Save address list failed (name={AddressListName}): {result.Error?.Message}", "MailFlow");
                return;
            }

            await RefreshAllAsync(ct);
        }, cancellationToken);
    }

    private async Task RemoveAddressListAsync(CancellationToken cancellationToken)
    {
        if (SelectedAddressList == null)
        {
            return;
        }

        var confirmed = ErrorDialogService.ShowConfirmation("Confirm tenant-wide deletion", $"Operation: Deleting address list\nTarget: {SelectedAddressList.DisplayName}\nImpact: can affect segmented GALs and address book policies.\n\nConfirm?");
        if (!confirmed)
        {
            return;
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.RemoveAddressListAsync(new RemoveAddressListRequest
            {
                Identity = SelectedAddressList.Identity
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                var error = result.Error?.Message ?? "Unable to delete address list";
                SetError(error);
                ShellViewModel.AddLog(LogLevel.Error, $"Remove address list failed (name={SelectedAddressList.Name}): {error}", "MailFlow");
                return;
            }

            await RefreshAllAsync(ct);
        }, cancellationToken);
    }

    private void ResetEditor()
    {
        AddressListIdentity = null;
        AddressListName = string.Empty;
        AddressListDisplayName = string.Empty;
        AddressListRecipientFilter = string.Empty;
        AddressListRecipientContainer = string.Empty;
        AddressListIncludedRecipients = string.Empty;
        AddressListConditionalCompany = string.Empty;
        AddressListConditionalDepartment = string.Empty;
        AddressListConditionalStateOrProvince = string.Empty;
        AddressListConditionalCustomAttribute1 = string.Empty;
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

