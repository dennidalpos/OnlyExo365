using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

internal sealed class MailFlowAddressBookPoliciesViewModel : MailFlowSectionViewModelBase
{
    private static readonly string[] ValidationProperties =
    {
        nameof(IsAddressBookPolicyInputValid),
        nameof(AddressBookPolicyValidationMessage)
    };

    private AddressBookPolicyDto? _selectedAddressBookPolicy;
    private string? _addressBookPolicyIdentity;
    private string _addressBookPolicyName = string.Empty;
    private string _addressBookPolicyAddressLists = string.Empty;
    private string _addressBookPolicyGlobalAddressList = string.Empty;
    private string _addressBookPolicyOfflineAddressBook = string.Empty;
    private string _addressBookPolicyRoomList = string.Empty;

    public MailFlowAddressBookPoliciesViewModel(
        IMailFlowWorkerService workerService,
        ShellViewModel shellViewModel,
        MailFlowOperationCoordinator coordinator,
        Func<CancellationToken, Task> refreshAllAsync)
        : base(workerService, shellViewModel, coordinator, refreshAllAsync)
    {
        NewAddressBookPolicyCommand = new AsyncRelayCommand(NewAddressBookPolicyAsync, () => !Coordinator.IsLoading);
        SaveAddressBookPolicyCommand = new AsyncRelayCommand(SaveAddressBookPolicyAsync, () => !Coordinator.IsLoading && IsAddressBookPolicyInputValid);
        RemoveAddressBookPolicyCommand = new AsyncRelayCommand(RemoveAddressBookPolicyAsync, () => !Coordinator.IsLoading && SelectedAddressBookPolicy != null);
    }

    public ObservableCollection<AddressBookPolicyDto> AddressBookPolicies { get; } = new();

    public AddressBookPolicyDto? SelectedAddressBookPolicy
    {
        get => _selectedAddressBookPolicy;
        set
        {
            if (SetProperty(ref _selectedAddressBookPolicy, value))
            {
                if (value != null)
                {
                    AddressBookPolicyIdentity = value.Identity;
                    AddressBookPolicyName = value.Name;
                    AddressBookPolicyAddressLists = string.Join(",", value.AddressLists);
                    AddressBookPolicyGlobalAddressList = value.GlobalAddressList;
                    AddressBookPolicyOfflineAddressBook = value.OfflineAddressBook;
                    AddressBookPolicyRoomList = value.RoomList;
                }
                else
                {
                    ResetEditor();
                }

                OnPropertyChanged(nameof(CanEditSelectedAddressBookPolicy));
                InvalidateCommands();
            }
        }
    }

    public bool CanEditSelectedAddressBookPolicy => SelectedAddressBookPolicy != null && !Coordinator.IsLoading;
    public bool IsAddressBookPolicyInputValid =>
        !string.IsNullOrWhiteSpace(AddressBookPolicyName) &&
        MailFlowViewModelSupport.SplitCsv(AddressBookPolicyAddressLists).Count > 0 &&
        !string.IsNullOrWhiteSpace(AddressBookPolicyGlobalAddressList) &&
        !string.IsNullOrWhiteSpace(AddressBookPolicyOfflineAddressBook);

    public string AddressBookPolicyValidationMessage
        => IsAddressBookPolicyInputValid
            ? string.Empty
            : "Address book policy: Name, at least one AddressList, a GlobalAddressList, and an OfflineAddressBook are required.";

    public string? AddressBookPolicyIdentity { get => _addressBookPolicyIdentity; set => SetEditorProperty(ref _addressBookPolicyIdentity, value); }
    public string AddressBookPolicyName { get => _addressBookPolicyName; set => SetEditorProperty(ref _addressBookPolicyName, value); }
    public string AddressBookPolicyAddressLists { get => _addressBookPolicyAddressLists; set => SetEditorProperty(ref _addressBookPolicyAddressLists, value); }
    public string AddressBookPolicyGlobalAddressList { get => _addressBookPolicyGlobalAddressList; set => SetEditorProperty(ref _addressBookPolicyGlobalAddressList, value); }
    public string AddressBookPolicyOfflineAddressBook { get => _addressBookPolicyOfflineAddressBook; set => SetEditorProperty(ref _addressBookPolicyOfflineAddressBook, value); }
    public string AddressBookPolicyRoomList { get => _addressBookPolicyRoomList; set => SetEditorProperty(ref _addressBookPolicyRoomList, value); }

    public ICommand NewAddressBookPolicyCommand { get; }
    public ICommand SaveAddressBookPolicyCommand { get; }
    public ICommand RemoveAddressBookPolicyCommand { get; }

    public async Task LoadAsync(CancellationToken cancellationToken)
    {
        var result = await WorkerService.GetAddressBookPoliciesAsync(new GetAddressBookPoliciesRequest(), cancellationToken: cancellationToken);
        if (!result.IsSuccess)
        {
            var error = result.Error?.Message ?? "Unable to load address book policies";
            SetError(error);
            ShellViewModel.AddLog(LogLevel.Error, $"MailFlow address book policies load failed: {error}", "MailFlow");
            return;
        }

        AddressBookPolicies.Clear();
        foreach (var item in result.Value?.Policies ?? new List<AddressBookPolicyDto>())
        {
            AddressBookPolicies.Add(item);
        }
    }

    private Task NewAddressBookPolicyAsync(CancellationToken cancellationToken)
    {
        SelectedAddressBookPolicy = null;
        ResetEditor();
        SetError(null);
        return Task.CompletedTask;
    }

    private async Task SaveAddressBookPolicyAsync(CancellationToken cancellationToken)
    {
        if (!IsAddressBookPolicyInputValid)
        {
            return;
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.UpsertAddressBookPolicyAsync(new UpsertAddressBookPolicyRequest
            {
                Identity = string.IsNullOrWhiteSpace(AddressBookPolicyIdentity) ? null : AddressBookPolicyIdentity,
                Name = AddressBookPolicyName.Trim(),
                AddressLists = MailFlowViewModelSupport.SplitCsv(AddressBookPolicyAddressLists),
                GlobalAddressList = AddressBookPolicyGlobalAddressList.Trim(),
                OfflineAddressBook = AddressBookPolicyOfflineAddressBook.Trim(),
                RoomList = AddressBookPolicyRoomList.Trim()
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                var correlationId = Guid.NewGuid().ToString("N");
                SetError($"Saving address book policy failed (ref: {correlationId}).");
                ShellViewModel.AddLog(LogLevel.Error, $"[{correlationId}] Save address book policy failed (name={AddressBookPolicyName}): {result.Error?.Message}", "MailFlow");
                return;
            }

            await RefreshAllAsync(ct);
        }, cancellationToken);
    }

    private async Task RemoveAddressBookPolicyAsync(CancellationToken cancellationToken)
    {
        if (SelectedAddressBookPolicy == null)
        {
            return;
        }

        var confirmed = ErrorDialogService.ShowConfirmation("Confirm tenant-wide deletion", $"Operation: Deleting address book policy\nTarget: {SelectedAddressBookPolicy.Name}\nImpact: can change address book segmentation for assigned mailboxes.\n\nConfirm?");
        if (!confirmed)
        {
            return;
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.RemoveAddressBookPolicyAsync(new RemoveAddressBookPolicyRequest
            {
                Identity = SelectedAddressBookPolicy.Identity
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                var error = result.Error?.Message ?? "Unable to delete address book policy";
                SetError(error);
                ShellViewModel.AddLog(LogLevel.Error, $"Remove address book policy failed (name={SelectedAddressBookPolicy.Name}): {error}", "MailFlow");
                return;
            }

            await RefreshAllAsync(ct);
        }, cancellationToken);
    }

    private void ResetEditor()
    {
        AddressBookPolicyIdentity = null;
        AddressBookPolicyName = string.Empty;
        AddressBookPolicyAddressLists = string.Empty;
        AddressBookPolicyGlobalAddressList = string.Empty;
        AddressBookPolicyOfflineAddressBook = string.Empty;
        AddressBookPolicyRoomList = string.Empty;
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
