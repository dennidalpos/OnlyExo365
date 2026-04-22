using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;

namespace OnlyExo365.Shell.ViewModels;

public sealed class ResourceMailboxEditorViewModel : ViewModelBase
{
    private bool _isInitializing;
    private bool _hasPendingChanges;
    private string? _resourceIdentity;
    private string _resourceType = "Room";
    private string _displayName = string.Empty;
    private string _name = string.Empty;
    private string _alias = string.Empty;
    private string _primarySmtpAddress = string.Empty;
    private bool _hiddenFromAddressListsEnabled;
    private string _fullAccessUsers = string.Empty;
    private string _sendAsUsers = string.Empty;
    private string _sendOnBehalfUsers = string.Empty;
    private HashSet<string> _originalFullAccessUsers = new(StringComparer.OrdinalIgnoreCase);
    private HashSet<string> _originalSendAsUsers = new(StringComparer.OrdinalIgnoreCase);
    private HashSet<string> _originalSendOnBehalfUsers = new(StringComparer.OrdinalIgnoreCase);
    private ResourceMailboxEditorSnapshot _originalState = ResourceMailboxEditorSnapshot.Default;

    public ResourceMailboxEditorViewModel()
    {
        BookingSettings = new ResourceBookingSettingsEditorViewModel();
        BookingSettings.PropertyChanged += OnBookingSettingsPropertyChanged;
        BeginCreate();
    }

    public ResourceBookingSettingsEditorViewModel BookingSettings { get; }

    public IReadOnlyList<string> ResourceTypes { get; } = ["Room", "Equipment"];

    public IReadOnlyList<string> AutomateProcessingModes { get; } = ["AutoAccept", "AutoUpdate", "None"];

    public bool HasPendingChanges
    {
        get => _hasPendingChanges;
        private set => SetProperty(ref _hasPendingChanges, value);
    }

    public string? ResourceIdentity
    {
        get => _resourceIdentity;
        private set
        {
            if (SetProperty(ref _resourceIdentity, value))
            {
                OnPropertyChanged(nameof(IsEditingExistingResource));
            }
        }
    }

    public bool IsEditingExistingResource => !string.IsNullOrWhiteSpace(ResourceIdentity);

    public string ResourceType
    {
        get => _resourceType;
        set => SetTrackedProperty(ref _resourceType, value, nameof(CanSave));
    }

    public string DisplayName
    {
        get => _displayName;
        set => SetTrackedProperty(ref _displayName, value, nameof(CanSave));
    }

    public string Name
    {
        get => _name;
        set => SetTrackedProperty(ref _name, value);
    }

    public string Alias
    {
        get => _alias;
        set => SetTrackedProperty(ref _alias, value, nameof(CanSave));
    }

    public string PrimarySmtpAddress
    {
        get => _primarySmtpAddress;
        set => SetTrackedProperty(ref _primarySmtpAddress, value, nameof(CanSave));
    }

    public bool HiddenFromAddressListsEnabled
    {
        get => _hiddenFromAddressListsEnabled;
        set => SetTrackedProperty(ref _hiddenFromAddressListsEnabled, value);
    }

    public string FullAccessUsers
    {
        get => _fullAccessUsers;
        set => SetTrackedProperty(ref _fullAccessUsers, value);
    }

    public string SendAsUsers
    {
        get => _sendAsUsers;
        set => SetTrackedProperty(ref _sendAsUsers, value);
    }

    public string SendOnBehalfUsers
    {
        get => _sendOnBehalfUsers;
        set => SetTrackedProperty(ref _sendOnBehalfUsers, value);
    }

    public bool CanSave =>
        !string.IsNullOrWhiteSpace(DisplayName) &&
        !string.IsNullOrWhiteSpace(Alias) &&
        !string.IsNullOrWhiteSpace(PrimarySmtpAddress);

    public void BeginCreate()
    {
        _isInitializing = true;

        ResourceIdentity = null;
        ResourceType = "Room";
        DisplayName = string.Empty;
        Name = string.Empty;
        Alias = string.Empty;
        PrimarySmtpAddress = string.Empty;
        HiddenFromAddressListsEnabled = false;
        FullAccessUsers = string.Empty;
        SendAsUsers = string.Empty;
        SendOnBehalfUsers = string.Empty;
        _originalFullAccessUsers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        _originalSendAsUsers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        _originalSendOnBehalfUsers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        BookingSettings.Reset();

        _isInitializing = false;
        AcceptCurrentStateAsOriginal();
    }

    public void ApplyDetails(ResourceMailboxDetailsDto details)
    {
        _isInitializing = true;

        ResourceIdentity = details.Identity;
        ResourceType = string.IsNullOrWhiteSpace(details.ResourceType) ? "Room" : details.ResourceType;
        DisplayName = details.DisplayName;
        Name = details.Name ?? details.DisplayName;
        Alias = details.Alias;
        PrimarySmtpAddress = details.PrimarySmtpAddress;
        HiddenFromAddressListsEnabled = details.HiddenFromAddressListsEnabled;

        _originalFullAccessUsers = ResourcePermissionDeltaBuilder.GetOriginalFullAccessUsers(details.Permissions);
        _originalSendAsUsers = ResourcePermissionDeltaBuilder.GetOriginalSendAsUsers(details.Permissions);
        _originalSendOnBehalfUsers = ResourcePermissionDeltaBuilder.GetOriginalSendOnBehalfUsers(details.Permissions);

        FullAccessUsers = ResourceCsvHelper.ToCsv(_originalFullAccessUsers);
        SendAsUsers = ResourceCsvHelper.ToCsv(_originalSendAsUsers);
        SendOnBehalfUsers = ResourceCsvHelper.ToCsv(_originalSendOnBehalfUsers);

        BookingSettings.Apply(details.BookingSettings);

        _isInitializing = false;
        AcceptCurrentStateAsOriginal();
    }

    public UpsertResourceMailboxRequest BuildUpsertRequest()
    {
        return new UpsertResourceMailboxRequest
        {
            Identity = ResourceIdentity,
            ResourceType = ResourceType,
            DisplayName = DisplayName.Trim(),
            Name = string.IsNullOrWhiteSpace(Name) ? DisplayName.Trim() : Name.Trim(),
            Alias = Alias.Trim(),
            PrimarySmtpAddress = PrimarySmtpAddress.Trim(),
            HiddenFromAddressListsEnabled = HiddenFromAddressListsEnabled,
            BookingSettings = BookingSettings.BuildDto()
        };
    }

    public List<PermissionDeltaActionDto> BuildPermissionDelta()
    {
        return ResourcePermissionDeltaBuilder.Build(
            FullAccessUsers,
            _originalFullAccessUsers,
            SendAsUsers,
            _originalSendAsUsers,
            SendOnBehalfUsers,
            _originalSendOnBehalfUsers);
    }

    public void AcceptCurrentStateAsOriginal(string? identity = null)
    {
        if (!string.IsNullOrWhiteSpace(identity))
        {
            ResourceIdentity = identity;
        }

        _originalFullAccessUsers = ResourceCsvHelper.ToSet(FullAccessUsers);
        _originalSendAsUsers = ResourceCsvHelper.ToSet(SendAsUsers);
        _originalSendOnBehalfUsers = ResourceCsvHelper.ToSet(SendOnBehalfUsers);
        _originalState = CaptureSnapshot();
        BookingSettings.AcceptCurrentStateAsOriginal();
        HasPendingChanges = false;
    }

    private void SetTrackedProperty<T>(ref T field, T value, params string[] additionalProperties)
    {
        if (!SetProperty(ref field, value))
        {
            return;
        }

        foreach (var property in additionalProperties)
        {
            OnPropertyChanged(property);
        }

        UpdatePendingChanges();
    }

    private void OnBookingSettingsPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ResourceBookingSettingsEditorViewModel.HasPendingChanges))
        {
            UpdatePendingChanges();
        }
    }

    private void UpdatePendingChanges()
    {
        if (_isInitializing)
        {
            return;
        }

        HasPendingChanges = !_originalState.Equals(CaptureSnapshot()) || BookingSettings.HasPendingChanges;
    }

    private ResourceMailboxEditorSnapshot CaptureSnapshot()
    {
        return new ResourceMailboxEditorSnapshot(
            NormalizeInput(ResourceIdentity),
            NormalizeInput(ResourceType),
            NormalizeInput(DisplayName),
            NormalizeInput(Name),
            NormalizeInput(Alias),
            NormalizeInput(PrimarySmtpAddress),
            HiddenFromAddressListsEnabled,
            ResourceCsvHelper.NormalizeCsv(FullAccessUsers),
            ResourceCsvHelper.NormalizeCsv(SendAsUsers),
            ResourceCsvHelper.NormalizeCsv(SendOnBehalfUsers));
    }

    private static string NormalizeInput(string? value) => string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();

    private sealed record ResourceMailboxEditorSnapshot(
        string ResourceIdentity,
        string ResourceType,
        string DisplayName,
        string Name,
        string Alias,
        string PrimarySmtpAddress,
        bool HiddenFromAddressListsEnabled,
        string FullAccessUsers,
        string SendAsUsers,
        string SendOnBehalfUsers)
    {
        public static ResourceMailboxEditorSnapshot Default { get; } = new(
            string.Empty,
            "Room",
            string.Empty,
            string.Empty,
            string.Empty,
            string.Empty,
            false,
            string.Empty,
            string.Empty,
            string.Empty);
    }
}

