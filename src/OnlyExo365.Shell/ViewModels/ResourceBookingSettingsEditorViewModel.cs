using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Shell.ViewModels;

public sealed class ResourceBookingSettingsEditorViewModel : ViewModelBase
{
    private bool _isInitializing;
    private bool _hasPendingChanges;
    private string _automateProcessing = "AutoAccept";
    private bool _allowConflicts;
    private bool _allBookInPolicy = true;
    private bool _allRequestInPolicy;
    private bool _allRequestOutOfPolicy;
    private int _bookingWindowInDays = 180;
    private int _maximumDurationInMinutes = 1440;
    private bool _deleteSubject = true;
    private bool _addOrganizerToSubject;
    private bool _removePrivateProperty = true;
    private bool _enforceSchedulingHorizon = true;
    private string _bookInPolicy = string.Empty;
    private string _requestInPolicy = string.Empty;
    private string _requestOutOfPolicy = string.Empty;
    private string _resourceDelegates = string.Empty;
    private ResourceBookingSettingsSnapshot _originalSettings = ResourceBookingSettingsSnapshot.Default;

    public bool HasPendingChanges
    {
        get => _hasPendingChanges;
        private set => SetProperty(ref _hasPendingChanges, value);
    }

    public string AutomateProcessing
    {
        get => _automateProcessing;
        set => SetTrackedProperty(ref _automateProcessing, value);
    }

    public bool AllowConflicts
    {
        get => _allowConflicts;
        set => SetTrackedProperty(ref _allowConflicts, value);
    }

    public bool AllBookInPolicy
    {
        get => _allBookInPolicy;
        set => SetTrackedProperty(ref _allBookInPolicy, value);
    }

    public bool AllRequestInPolicy
    {
        get => _allRequestInPolicy;
        set => SetTrackedProperty(ref _allRequestInPolicy, value);
    }

    public bool AllRequestOutOfPolicy
    {
        get => _allRequestOutOfPolicy;
        set => SetTrackedProperty(ref _allRequestOutOfPolicy, value);
    }

    public int BookingWindowInDays
    {
        get => _bookingWindowInDays;
        set => SetTrackedProperty(ref _bookingWindowInDays, value);
    }

    public int MaximumDurationInMinutes
    {
        get => _maximumDurationInMinutes;
        set => SetTrackedProperty(ref _maximumDurationInMinutes, value);
    }

    public bool DeleteSubject
    {
        get => _deleteSubject;
        set => SetTrackedProperty(ref _deleteSubject, value);
    }

    public bool AddOrganizerToSubject
    {
        get => _addOrganizerToSubject;
        set => SetTrackedProperty(ref _addOrganizerToSubject, value);
    }

    public bool RemovePrivateProperty
    {
        get => _removePrivateProperty;
        set => SetTrackedProperty(ref _removePrivateProperty, value);
    }

    public bool EnforceSchedulingHorizon
    {
        get => _enforceSchedulingHorizon;
        set => SetTrackedProperty(ref _enforceSchedulingHorizon, value);
    }

    public string BookInPolicy
    {
        get => _bookInPolicy;
        set => SetTrackedProperty(ref _bookInPolicy, value);
    }

    public string RequestInPolicy
    {
        get => _requestInPolicy;
        set => SetTrackedProperty(ref _requestInPolicy, value);
    }

    public string RequestOutOfPolicy
    {
        get => _requestOutOfPolicy;
        set => SetTrackedProperty(ref _requestOutOfPolicy, value);
    }

    public string ResourceDelegates
    {
        get => _resourceDelegates;
        set => SetTrackedProperty(ref _resourceDelegates, value);
    }

    public void Reset()
    {
        _isInitializing = true;

        AutomateProcessing = "AutoAccept";
        AllowConflicts = false;
        AllBookInPolicy = true;
        AllRequestInPolicy = false;
        AllRequestOutOfPolicy = false;
        BookingWindowInDays = 180;
        MaximumDurationInMinutes = 1440;
        DeleteSubject = true;
        AddOrganizerToSubject = false;
        RemovePrivateProperty = true;
        EnforceSchedulingHorizon = true;
        BookInPolicy = string.Empty;
        RequestInPolicy = string.Empty;
        RequestOutOfPolicy = string.Empty;
        ResourceDelegates = string.Empty;

        _isInitializing = false;
        AcceptCurrentStateAsOriginal();
    }

    public void Apply(ResourceBookingSettingsDto? bookingSettings)
    {
        _isInitializing = true;

        var booking = bookingSettings ?? new ResourceBookingSettingsDto();
        AutomateProcessing = string.IsNullOrWhiteSpace(booking.AutomateProcessing) ? "AutoAccept" : booking.AutomateProcessing;
        AllowConflicts = booking.AllowConflicts;
        AllBookInPolicy = booking.AllBookInPolicy;
        AllRequestInPolicy = booking.AllRequestInPolicy;
        AllRequestOutOfPolicy = booking.AllRequestOutOfPolicy;
        BookingWindowInDays = booking.BookingWindowInDays ?? 180;
        MaximumDurationInMinutes = booking.MaximumDurationInMinutes ?? 1440;
        DeleteSubject = booking.DeleteSubject ?? true;
        AddOrganizerToSubject = booking.AddOrganizerToSubject ?? false;
        RemovePrivateProperty = booking.RemovePrivateProperty ?? true;
        EnforceSchedulingHorizon = booking.EnforceSchedulingHorizon ?? true;
        BookInPolicy = ResourceCsvHelper.ToCsv(booking.BookInPolicy);
        RequestInPolicy = ResourceCsvHelper.ToCsv(booking.RequestInPolicy);
        RequestOutOfPolicy = ResourceCsvHelper.ToCsv(booking.RequestOutOfPolicy);
        ResourceDelegates = ResourceCsvHelper.ToCsv(booking.ResourceDelegates);

        _isInitializing = false;
        AcceptCurrentStateAsOriginal();
    }

    public ResourceBookingSettingsDto BuildDto()
    {
        return new ResourceBookingSettingsDto
        {
            AutomateProcessing = AutomateProcessing,
            AllowConflicts = AllowConflicts,
            AllBookInPolicy = AllBookInPolicy,
            AllRequestInPolicy = AllRequestInPolicy,
            AllRequestOutOfPolicy = AllRequestOutOfPolicy,
            BookingWindowInDays = BookingWindowInDays,
            MaximumDurationInMinutes = MaximumDurationInMinutes,
            DeleteSubject = DeleteSubject,
            AddOrganizerToSubject = AddOrganizerToSubject,
            RemovePrivateProperty = RemovePrivateProperty,
            EnforceSchedulingHorizon = EnforceSchedulingHorizon,
            BookInPolicy = ResourceCsvHelper.Parse(BookInPolicy),
            RequestInPolicy = ResourceCsvHelper.Parse(RequestInPolicy),
            RequestOutOfPolicy = ResourceCsvHelper.Parse(RequestOutOfPolicy),
            ResourceDelegates = ResourceCsvHelper.Parse(ResourceDelegates)
        };
    }

    public void AcceptCurrentStateAsOriginal()
    {
        _originalSettings = CaptureSnapshot();
        HasPendingChanges = false;
    }

    private void SetTrackedProperty<T>(ref T field, T value)
    {
        if (!SetProperty(ref field, value))
        {
            return;
        }

        UpdatePendingChanges();
    }

    private void UpdatePendingChanges()
    {
        if (_isInitializing)
        {
            return;
        }

        HasPendingChanges = !_originalSettings.Equals(CaptureSnapshot());
    }

    private ResourceBookingSettingsSnapshot CaptureSnapshot()
    {
        return new ResourceBookingSettingsSnapshot(
            AutomateProcessing,
            AllowConflicts,
            AllBookInPolicy,
            AllRequestInPolicy,
            AllRequestOutOfPolicy,
            BookingWindowInDays,
            MaximumDurationInMinutes,
            DeleteSubject,
            AddOrganizerToSubject,
            RemovePrivateProperty,
            EnforceSchedulingHorizon,
            ResourceCsvHelper.NormalizeCsv(BookInPolicy),
            ResourceCsvHelper.NormalizeCsv(RequestInPolicy),
            ResourceCsvHelper.NormalizeCsv(RequestOutOfPolicy),
            ResourceCsvHelper.NormalizeCsv(ResourceDelegates));
    }

    private sealed record ResourceBookingSettingsSnapshot(
        string AutomateProcessing,
        bool AllowConflicts,
        bool AllBookInPolicy,
        bool AllRequestInPolicy,
        bool AllRequestOutOfPolicy,
        int BookingWindowInDays,
        int MaximumDurationInMinutes,
        bool DeleteSubject,
        bool AddOrganizerToSubject,
        bool RemovePrivateProperty,
        bool EnforceSchedulingHorizon,
        string BookInPolicy,
        string RequestInPolicy,
        string RequestOutOfPolicy,
        string ResourceDelegates)
    {
        public static ResourceBookingSettingsSnapshot Default { get; } = new(
            "AutoAccept",
            false,
            true,
            false,
            false,
            180,
            1440,
            true,
            false,
            true,
            true,
            string.Empty,
            string.Empty,
            string.Empty,
            string.Empty);
    }
}

