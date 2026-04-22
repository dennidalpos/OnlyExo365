using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Shell.Helpers;

namespace OnlyExo365.Shell.ViewModels;

internal sealed class MailFlowOfflineAddressBooksViewModel : MailFlowSectionViewModelBase
{
    private static readonly string[] ValidationProperties =
    {
        nameof(IsOfflineAddressBookInputValid),
        nameof(OfflineAddressBookValidationMessage)
    };

    private OfflineAddressBookDto? _selectedOfflineAddressBook;
    private string? _offlineAddressBookIdentity;
    private string _offlineAddressBookName = string.Empty;
    private string _offlineAddressBookAddressLists = string.Empty;
    private string _offlineAddressBookDiffRetentionPeriod = string.Empty;
    private bool _offlineAddressBookIsDefault;

    public MailFlowOfflineAddressBooksViewModel(
        IMailFlowWorkerService workerService,
        ShellViewModel shellViewModel,
        MailFlowOperationCoordinator coordinator,
        Func<CancellationToken, Task> refreshAllAsync)
        : base(workerService, shellViewModel, coordinator, refreshAllAsync)
    {
        NewOfflineAddressBookCommand = new AsyncRelayCommand(NewOfflineAddressBookAsync, () => IsSectionSupported && !Coordinator.IsLoading);
        SaveOfflineAddressBookCommand = new AsyncRelayCommand(SaveOfflineAddressBookAsync, () => IsSectionSupported && !Coordinator.IsLoading && IsOfflineAddressBookInputValid);
        RemoveOfflineAddressBookCommand = new AsyncRelayCommand(RemoveOfflineAddressBookAsync, () => IsSectionSupported && !Coordinator.IsLoading && SelectedOfflineAddressBook != null && !SelectedOfflineAddressBook.IsDefault);
    }

    public ObservableCollection<OfflineAddressBookDto> OfflineAddressBooks { get; } = new();

    public OfflineAddressBookDto? SelectedOfflineAddressBook
    {
        get => _selectedOfflineAddressBook;
        set
        {
            if (SetProperty(ref _selectedOfflineAddressBook, value))
            {
                if (value != null)
                {
                    OfflineAddressBookIdentity = value.Identity;
                    OfflineAddressBookName = value.Name;
                    OfflineAddressBookAddressLists = string.Join(",", value.AddressLists);
                    OfflineAddressBookDiffRetentionPeriod = value.DiffRetentionPeriod?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
                    OfflineAddressBookIsDefault = value.IsDefault;
                }
                else
                {
                    ResetEditor();
                }

                OnPropertyChanged(nameof(CanEditSelectedOfflineAddressBook));
                InvalidateCommands();
            }
        }
    }

    public bool CanEditSelectedOfflineAddressBook => SelectedOfflineAddressBook != null && IsSectionSupported && !Coordinator.IsLoading;
    public bool IsOfflineAddressBookInputValid =>
        !string.IsNullOrWhiteSpace(OfflineAddressBookName) &&
        MailFlowViewModelSupport.SplitCsv(OfflineAddressBookAddressLists).Count > 0 &&
        (string.IsNullOrWhiteSpace(OfflineAddressBookDiffRetentionPeriod) ||
         int.TryParse(OfflineAddressBookDiffRetentionPeriod, out var diffRetention) && diffRetention >= 0);

    public string OfflineAddressBookValidationMessage
        => IsOfflineAddressBookInputValid
            ? string.Empty
            : "Offline address book: Name is required, at least one AddressList must be provided, and DiffRetentionPeriod is optional but must be a non-negative integer.";

    public string? OfflineAddressBookIdentity { get => _offlineAddressBookIdentity; set => SetEditorProperty(ref _offlineAddressBookIdentity, value); }
    public string OfflineAddressBookName { get => _offlineAddressBookName; set => SetEditorProperty(ref _offlineAddressBookName, value); }
    public string OfflineAddressBookAddressLists { get => _offlineAddressBookAddressLists; set => SetEditorProperty(ref _offlineAddressBookAddressLists, value); }
    public string OfflineAddressBookDiffRetentionPeriod { get => _offlineAddressBookDiffRetentionPeriod; set => SetEditorProperty(ref _offlineAddressBookDiffRetentionPeriod, value); }
    public bool OfflineAddressBookIsDefault { get => _offlineAddressBookIsDefault; set => SetProperty(ref _offlineAddressBookIsDefault, value); }

    public ICommand NewOfflineAddressBookCommand { get; }
    public ICommand SaveOfflineAddressBookCommand { get; }
    public ICommand RemoveOfflineAddressBookCommand { get; }

    public async Task LoadAsync(CancellationToken cancellationToken)
    {
        var result = await WorkerService.GetOfflineAddressBooksAsync(new GetOfflineAddressBooksRequest(), cancellationToken: cancellationToken);
        if (!result.IsSuccess)
        {
            var error = result.Error?.Message ?? "Unable to load offline address books";
            SetError(error);
            ShellViewModel.AddLog(LogLevel.Error, $"MailFlow offline address books load failed: {error}", "MailFlow");
            return;
        }

        OfflineAddressBooks.Clear();
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
            SelectedOfflineAddressBook = null;
            return;
        }

        foreach (var item in result.Value?.OfflineAddressBooks ?? new List<OfflineAddressBookDto>())
        {
            OfflineAddressBooks.Add(item);
        }
    }

    private Task NewOfflineAddressBookAsync(CancellationToken cancellationToken)
    {
        SelectedOfflineAddressBook = null;
        ResetEditor();
        SetError(null);
        return Task.CompletedTask;
    }

    private async Task SaveOfflineAddressBookAsync(CancellationToken cancellationToken)
    {
        if (!IsOfflineAddressBookInputValid)
        {
            return;
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.UpsertOfflineAddressBookAsync(new UpsertOfflineAddressBookRequest
            {
                Identity = string.IsNullOrWhiteSpace(OfflineAddressBookIdentity) ? null : OfflineAddressBookIdentity,
                Name = OfflineAddressBookName.Trim(),
                AddressLists = MailFlowViewModelSupport.SplitCsv(OfflineAddressBookAddressLists),
                DiffRetentionPeriod = int.TryParse(OfflineAddressBookDiffRetentionPeriod, out var diffRetention) ? diffRetention : null
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                var correlationId = Guid.NewGuid().ToString("N");
                SetError($"Saving offline address book failed (ref: {correlationId}).");
                ShellViewModel.AddLog(LogLevel.Error, $"[{correlationId}] Save offline address book failed (name={OfflineAddressBookName}): {result.Error?.Message}", "MailFlow");
                return;
            }

            await RefreshAllAsync(ct);
        }, cancellationToken);
    }

    private async Task RemoveOfflineAddressBookAsync(CancellationToken cancellationToken)
    {
        if (SelectedOfflineAddressBook == null || SelectedOfflineAddressBook.IsDefault)
        {
            return;
        }

        var confirmed = ErrorDialogService.ShowConfirmation("Confirm tenant-wide deletion", $"Operation: Deleting offline address book\nTarget: {SelectedOfflineAddressBook.Name}\nImpact: can affect OAB distribution for legacy or segmented clients.\n\nConfirm?");
        if (!confirmed)
        {
            return;
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.RemoveOfflineAddressBookAsync(new RemoveOfflineAddressBookRequest
            {
                Identity = SelectedOfflineAddressBook.Identity
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                var error = result.Error?.Message ?? "Unable to delete offline address book";
                SetError(error);
                ShellViewModel.AddLog(LogLevel.Error, $"Remove offline address book failed (name={SelectedOfflineAddressBook.Name}): {error}", "MailFlow");
                return;
            }

            await RefreshAllAsync(ct);
        }, cancellationToken);
    }

    private void ResetEditor()
    {
        OfflineAddressBookIdentity = null;
        OfflineAddressBookName = string.Empty;
        OfflineAddressBookAddressLists = string.Empty;
        OfflineAddressBookDiffRetentionPeriod = string.Empty;
        OfflineAddressBookIsDefault = false;
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

