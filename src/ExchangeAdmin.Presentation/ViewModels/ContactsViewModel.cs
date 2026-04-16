using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Paging;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Localization;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

public class ContactsViewModel : ViewModelBase
{
    private readonly IWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;
    private readonly DebounceHelper _searchDebounce = new();

    private bool _isLoading;
    private bool _isSaving;
    private string? _errorMessage;
    private string? _searchQuery;
    private string _contactKindFilter = "All";
    private int _totalCount;
    private int _currentSkip;
    private const int PageSize = PagingDefaults.DefaultPageSize;
    private bool _hasMore;
    private ContactListItemDto? _selectedContact;

    private string? _contactIdentity;
    private string _contactKind = "MailContact";
    private string _displayName = string.Empty;
    private string _name = string.Empty;
    private string _alias = string.Empty;
    private string _primarySmtpAddress = string.Empty;
    private string _externalEmailAddress = string.Empty;
    private string _userPrincipalName = string.Empty;
    private string? _pendingMailUserPassword;
    private int _mailUserPasswordClearTrigger;
    private bool _hiddenFromAddressListsEnabled;

    public ContactsViewModel(IWorkerService workerService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;

        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
        LoadMoreCommand = new AsyncRelayCommand(LoadMoreAsync, () => CanLoadMore);
        SaveCommand = new AsyncRelayCommand(SaveAsync, () => CanSave);
        RemoveCommand = new AsyncRelayCommand(RemoveAsync, () => CanRemove);
        NewContactCommand = new RelayCommand(BeginCreate, () => !IsSaving);

        _shellViewModel.PropertyChanged += OnShellPropertyChanged;
    }

    public ObservableCollection<ContactListItemDto> Contacts { get; } = new();
    public CollectionLoadProgressViewModel LoadProgress { get; } = new();

    public IReadOnlyList<string> ContactKinds { get; } = new[] { "MailContact", "MailUser" };
    public IReadOnlyList<string> ContactKindFilters { get; } = new[] { "All", "MailContact", "MailUser" };

    public bool IsLoading
    {
        get => _isLoading;
        private set
        {
            if (SetProperty(ref _isLoading, value))
            {
                OnPropertyChanged(nameof(CanRefresh));
                OnPropertyChanged(nameof(CanLoadMore));
                OnPropertyChanged(nameof(CanSave));
                OnPropertyChanged(nameof(CanRemove));
                CommandManager.InvalidateRequerySuggested();
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
                OnPropertyChanged(nameof(CanSave));
                OnPropertyChanged(nameof(CanRemove));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

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

    public string? SearchQuery
    {
        get => _searchQuery;
        set
        {
            if (SetProperty(ref _searchQuery, value))
            {
                _searchDebounce.Debounce(TriggerRefreshFromUi, 300);
            }
        }
    }

    public string ContactKindFilter
    {
        get => _contactKindFilter;
        set
        {
            if (SetProperty(ref _contactKindFilter, value))
            {
                TriggerRefreshFromUi();
            }
        }
    }

    public int TotalCount
    {
        get => _totalCount;
        private set
        {
            if (SetProperty(ref _totalCount, value))
            {
                OnPropertyChanged(nameof(StatusText));
            }
        }
    }

    public bool HasMore
    {
        get => _hasMore;
        private set
        {
            if (SetProperty(ref _hasMore, value))
            {
                OnPropertyChanged(nameof(CanLoadMore));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string StatusText => $"{Contacts.Count} of {TotalCount} contacts";

    public ContactListItemDto? SelectedContact
    {
        get => _selectedContact;
        set
        {
            if (SetProperty(ref _selectedContact, value) && value != null)
            {
                _ = LoadDetailsAsync(value);
            }
        }
    }

    public string? ContactIdentity
    {
        get => _contactIdentity;
        private set
        {
            if (SetProperty(ref _contactIdentity, value))
            {
                OnPropertyChanged(nameof(ContactIdentityDisplay));
                OnPropertyChanged(nameof(IsEditingExistingContact));
                OnPropertyChanged(nameof(IsCreateMailUser));
                OnPropertyChanged(nameof(CanSave));
                OnPropertyChanged(nameof(CanRemove));
            }
        }
    }

    public string? ContactIdentityDisplay => string.IsNullOrWhiteSpace(ContactIdentity)
        ? null
        : Loc.GetFormat("Contact.IdentityFormat", ContactIdentity);

    public bool IsEditingExistingContact => !string.IsNullOrWhiteSpace(ContactIdentity);
    public bool IsMailUser => string.Equals(ContactKind, "MailUser", StringComparison.OrdinalIgnoreCase);

    public string ContactKind
    {
        get => _contactKind;
        set
        {
            if (SetProperty(ref _contactKind, value))
            {
                if (!IsMailUser || IsEditingExistingContact)
                {
                    ClearMailUserPassword();
                }

                OnPropertyChanged(nameof(IsMailUser));
                OnPropertyChanged(nameof(IsCreateMailUser));
                OnPropertyChanged(nameof(CanSave));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool IsCreateMailUser => IsMailUser && !IsEditingExistingContact;

    public string DisplayName
    {
        get => _displayName;
        set
        {
            if (SetProperty(ref _displayName, value))
            {
                OnPropertyChanged(nameof(CanSave));
            }
        }
    }

    public string Name
    {
        get => _name;
        set => SetProperty(ref _name, value);
    }

    public string Alias
    {
        get => _alias;
        set
        {
            if (SetProperty(ref _alias, value))
            {
                OnPropertyChanged(nameof(CanSave));
            }
        }
    }

    public string PrimarySmtpAddress
    {
        get => _primarySmtpAddress;
        set
        {
            if (SetProperty(ref _primarySmtpAddress, value))
            {
                OnPropertyChanged(nameof(CanSave));
            }
        }
    }

    public string ExternalEmailAddress
    {
        get => _externalEmailAddress;
        set
        {
            if (SetProperty(ref _externalEmailAddress, value))
            {
                OnPropertyChanged(nameof(CanSave));
            }
        }
    }

    public string UserPrincipalName
    {
        get => _userPrincipalName;
        set
        {
            if (SetProperty(ref _userPrincipalName, value))
            {
                OnPropertyChanged(nameof(CanSave));
            }
        }
    }

    public bool HasMailUserPassword
    {
        get => !string.IsNullOrWhiteSpace(_pendingMailUserPassword);
    }

    public int MailUserPasswordClearTrigger
    {
        get => _mailUserPasswordClearTrigger;
        private set
        {
            SetProperty(ref _mailUserPasswordClearTrigger, value);
        }
    }

    public bool HiddenFromAddressListsEnabled
    {
        get => _hiddenFromAddressListsEnabled;
        set => SetProperty(ref _hiddenFromAddressListsEnabled, value);
    }

    public bool CanRefresh => !IsLoading && _shellViewModel.IsExchangeConnected;
    public bool CanLoadMore => !IsLoading && HasMore && _shellViewModel.IsExchangeConnected;
    public bool CanRemove => !IsSaving && IsEditingExistingContact;

    public bool CanSave =>
        !IsLoading &&
        !IsSaving &&
        _shellViewModel.IsExchangeConnected &&
        !string.IsNullOrWhiteSpace(DisplayName) &&
        !string.IsNullOrWhiteSpace(Alias) &&
        !string.IsNullOrWhiteSpace(PrimarySmtpAddress) &&
        !string.IsNullOrWhiteSpace(ExternalEmailAddress) &&
        (!IsMailUser || IsEditingExistingContact || (!string.IsNullOrWhiteSpace(UserPrincipalName) && HasMailUserPassword));

    public ICommand RefreshCommand { get; }
    public ICommand LoadMoreCommand { get; }
    public ICommand SaveCommand { get; }
    public ICommand RemoveCommand { get; }
    public ICommand NewContactCommand { get; }

    public async Task LoadAsync()
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            ClearMailUserPassword();
            Contacts.Clear();
            ErrorMessage = "Not connected to Exchange Online";
            return;
        }

        if (Contacts.Count == 0)
        {
            await RefreshAsync(CancellationToken.None);
        }
    }

    private void TriggerRefreshFromUi()
    {
        _ = SafeRefreshAsync();
    }

    private async Task SafeRefreshAsync()
    {
        try
        {
            await RefreshAsync(CancellationToken.None);
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Mail contacts refresh failed: {ex.Message}", "Contacts");
        }
    }

    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            ClearMailUserPassword();
            Contacts.Clear();
            ErrorMessage = "Not connected to Exchange Online";
            TotalCount = 0;
            HasMore = false;
            return;
        }

        IsLoading = true;
        LoadProgress.Start("Loading contacts...", "contacts");
        ErrorMessage = null;
        var refreshPageSize = GetRefreshPageSize(Contacts.Count);
        _currentSkip = 0;

        try
        {
            var result = await _workerService.GetContactsAsync(
                new GetContactsRequest
                {
                    ContactKind = NormalizeContactKindFilter(ContactKindFilter),
                    SearchQuery = string.IsNullOrWhiteSpace(SearchQuery) ? null : SearchQuery.Trim(),
                    PageSize = refreshPageSize,
                    Skip = 0
                },
                eventHandler: HandleLoadProgressEvent,
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to load contacts";
                return;
            }

            Contacts.ReplaceAll(result.Value.Contacts);
            TotalCount = result.Value.TotalCount;
            HasMore = result.Value.HasMore;
            _currentSkip = Contacts.Count;
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            LoadProgress.Reset();
            IsLoading = false;
        }
    }

    private async Task LoadMoreAsync(CancellationToken cancellationToken)
    {
        if (!HasMore)
        {
            return;
        }

        IsLoading = true;
        LoadProgress.Start("Loading contacts...", "contacts");
        ErrorMessage = null;

        try
        {
            var result = await _workerService.GetContactsAsync(
                new GetContactsRequest
                {
                    ContactKind = NormalizeContactKindFilter(ContactKindFilter),
                    SearchQuery = string.IsNullOrWhiteSpace(SearchQuery) ? null : SearchQuery.Trim(),
                    PageSize = PageSize,
                    Skip = _currentSkip
                },
                eventHandler: HandleLoadProgressEvent,
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to load contacts";
                return;
            }

            foreach (var item in result.Value.Contacts)
            {
                Contacts.Add(item);
            }

            TotalCount = result.Value.TotalCount;
            HasMore = result.Value.HasMore;
            _currentSkip = Contacts.Count;
            OnPropertyChanged(nameof(StatusText));
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            LoadProgress.Reset();
            IsLoading = false;
        }
    }

    private async Task LoadDetailsAsync(ContactListItemDto selected)
    {
        var selectedIdentity = selected.Identity;
        var selectedKind = selected.ContactKind;
        ErrorMessage = null;

        try
        {
            var result = await _workerService.GetContactDetailsAsync(
                new GetContactDetailsRequest
                {
                    Identity = selected.Identity,
                    ContactKind = selected.ContactKind
                },
                cancellationToken: CancellationToken.None);

            if (!result.IsSuccess || result.Value == null)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to load contact details.";
                return;
            }

            if (SelectedContact == null ||
                !string.Equals(SelectedContact.Identity, selectedIdentity, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(SelectedContact.ContactKind, selectedKind, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            ApplyDetails(selected, result.Value);
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
    }

    private async Task SaveAsync(CancellationToken cancellationToken)
    {
        if (!CanSave)
        {
            return;
        }

        if (!ConfirmMutation(
                string.IsNullOrWhiteSpace(ContactIdentity) ? "Creating contact" : "Updating contact",
                PrimarySmtpAddress.Trim(),
                $"Save the {ContactKind} contact.",
                "Confirm contact save"))
        {
            ClearMailUserPassword();
            return;
        }

        IsSaving = true;
        ErrorMessage = null;

        try
        {
            var password = IsCreateMailUser ? CaptureMailUserPassword() : null;
            var result = await _workerService.UpsertContactAsync(
                new UpsertContactRequest
                {
                    Identity = ContactIdentity,
                    ContactKind = ContactKind,
                    DisplayName = DisplayName.Trim(),
                    Name = string.IsNullOrWhiteSpace(Name) ? DisplayName.Trim() : Name.Trim(),
                    Alias = Alias.Trim(),
                    PrimarySmtpAddress = PrimarySmtpAddress.Trim(),
                    ExternalEmailAddress = ExternalEmailAddress.Trim(),
                    UserPrincipalName = string.IsNullOrWhiteSpace(UserPrincipalName) ? null : UserPrincipalName.Trim(),
                    Password = password,
                    HiddenFromAddressListsEnabled = HiddenFromAddressListsEnabled
                },
                cancellationToken: cancellationToken);

            ClearMailUserPassword();

            if (!result.IsSuccess)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to save the contact.";
                return;
            }

            _shellViewModel.AddLog(LogLevel.Information, $"Contact saved: {DisplayName} ({ContactKind})", "Contacts");
            await RefreshAsync(cancellationToken);
            await TryReselectSavedContactAsync(cancellationToken);
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

    private async Task RemoveAsync(CancellationToken cancellationToken)
    {
        if (!CanRemove || string.IsNullOrWhiteSpace(ContactIdentity))
        {
            return;
        }

        if (!ConfirmMutation(
                "Deleting contact",
                DisplayName,
                $"Permanently removes the {ContactKind} contact.",
                "Confirm contact deletion"))
        {
            return;
        }

        IsSaving = true;
        ErrorMessage = null;

        try
        {
            var result = await _workerService.RemoveContactAsync(
                new RemoveContactRequest
                {
                    Identity = ContactIdentity,
                    ContactKind = ContactKind
                },
                cancellationToken: cancellationToken);

            if (!result.IsSuccess)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to remove the contact.";
                return;
            }

            _shellViewModel.AddLog(LogLevel.Warning, $"Contact removed: {DisplayName} ({ContactKind})", "Contacts");
            BeginCreate();
            await RefreshAsync(cancellationToken);
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

    private void BeginCreate()
    {
        SelectedContact = null;
        ContactIdentity = null;
        ContactKind = "MailContact";
        DisplayName = string.Empty;
        Name = string.Empty;
        Alias = string.Empty;
        PrimarySmtpAddress = string.Empty;
        ExternalEmailAddress = string.Empty;
        UserPrincipalName = string.Empty;
        ClearMailUserPassword();
        HiddenFromAddressListsEnabled = false;
        ErrorMessage = null;
    }

    private void ApplyDetails(ContactListItemDto selected, ContactDetailsDto details)
    {
        var resolvedIdentity = FirstNonEmpty(details.Identity, selected.Identity);
        var resolvedKind = FirstNonEmpty(details.ContactKind, selected.ContactKind, "MailContact");
        var resolvedDisplayName = FirstNonEmpty(details.DisplayName, selected.DisplayName);
        var resolvedName = FirstNonEmpty(details.Name, selected.Name, resolvedDisplayName);
        var resolvedAlias = FirstNonEmpty(details.Alias, selected.Alias);
        var resolvedPrimarySmtpAddress = FirstNonEmpty(details.PrimarySmtpAddress, selected.PrimarySmtpAddress);
        var resolvedExternalEmailAddress = FirstNonEmpty(details.ExternalEmailAddress, selected.ExternalEmailAddress);
        var resolvedUserPrincipalName = FirstNonEmpty(details.UserPrincipalName, selected.UserPrincipalName);

        ContactIdentity = resolvedIdentity;
        ContactKind = resolvedKind;
        DisplayName = resolvedDisplayName;
        Name = resolvedName;
        Alias = resolvedAlias;
        PrimarySmtpAddress = resolvedPrimarySmtpAddress;
        ExternalEmailAddress = resolvedExternalEmailAddress;
        UserPrincipalName = resolvedUserPrincipalName;
        ClearMailUserPassword();
        HiddenFromAddressListsEnabled = details.HiddenFromAddressListsEnabled;
    }

    private async Task TryReselectSavedContactAsync(CancellationToken cancellationToken)
    {
        var match = Contacts.FirstOrDefault(contact =>
            string.Equals(contact.PrimarySmtpAddress, PrimarySmtpAddress, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(contact.ContactKind, ContactKind, StringComparison.OrdinalIgnoreCase));

        if (match != null)
        {
            SelectedContact = match;
            await LoadDetailsAsync(match);
        }
    }

    private static string? NormalizeContactKindFilter(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || string.Equals(value, "All", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return value;
    }

    public void SetMailUserPassword(string? value)
    {
        var normalized = string.IsNullOrEmpty(value) ? null : value;
        if (string.Equals(_pendingMailUserPassword, normalized, StringComparison.Ordinal))
        {
            return;
        }

        _pendingMailUserPassword = normalized;
        OnPropertyChanged(nameof(HasMailUserPassword));
        OnPropertyChanged(nameof(CanSave));
        CommandManager.InvalidateRequerySuggested();
    }

    private string? CaptureMailUserPassword()
    {
        return string.IsNullOrWhiteSpace(_pendingMailUserPassword)
            ? null
            : _pendingMailUserPassword.Trim();
    }

    private void ClearMailUserPassword()
    {
        var hadValue = !string.IsNullOrWhiteSpace(_pendingMailUserPassword);
        _pendingMailUserPassword = null;
        MailUserPasswordClearTrigger++;

        if (hadValue)
        {
            OnPropertyChanged(nameof(HasMailUserPassword));
            OnPropertyChanged(nameof(CanSave));
            CommandManager.InvalidateRequerySuggested();
        }
    }

    private void OnShellPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ShellViewModel.IsExchangeConnected) && !_shellViewModel.IsExchangeConnected)
        {
            ClearMailUserPassword();
        }
    }

    private static int GetRefreshPageSize(int loadedCount)
        => Math.Max(PageSize, loadedCount);

    private static string FirstNonEmpty(params string?[] values)
    {
        foreach (var value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return string.Empty;
    }

    private void HandleLoadProgressEvent(EventEnvelope evt)
    {
        _ = RunOnUiThreadAsync(() => LoadProgress.Apply(evt));
    }
}


