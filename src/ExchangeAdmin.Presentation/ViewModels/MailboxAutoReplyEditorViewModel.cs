using System.Net;
using System.Text.RegularExpressions;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Presentation.ViewModels;

public sealed class MailboxAutoReplyEditorViewModel : ViewModelBase
{
    private bool _hasPendingChanges;
    private bool _isInitializing;
    private MailboxAutoReplySnapshot? _originalSettings;
    private bool _autoReplyEnabled;
    private bool _autoReplyScheduled;
    private DateTime? _autoReplyStartDate;
    private string? _autoReplyStartTime;
    private DateTime? _autoReplyEndDate;
    private string? _autoReplyEndTime;
    private string? _autoReplyInternalMessage;
    private string? _autoReplyExternalMessage;
    private string? _autoReplyExternalAudience;

    public bool HasPendingChanges
    {
        get => _hasPendingChanges;
        private set => SetProperty(ref _hasPendingChanges, value);
    }

    public bool AutoReplyEnabled
    {
        get => _autoReplyEnabled;
        set
        {
            if (SetProperty(ref _autoReplyEnabled, value))
            {
                if (!value)
                {
                    AutoReplyScheduled = false;
                }

                UpdatePendingChanges();
                OnPropertyChanged(nameof(AutoReplyStateLabel));
            }
        }
    }

    public bool AutoReplyScheduled
    {
        get => _autoReplyScheduled;
        set
        {
            if (SetProperty(ref _autoReplyScheduled, value))
            {
                UpdatePendingChanges();
                OnPropertyChanged(nameof(AutoReplyStateLabel));
            }
        }
    }

    public DateTime? AutoReplyStartDate
    {
        get => _autoReplyStartDate;
        set
        {
            if (SetProperty(ref _autoReplyStartDate, value))
            {
                UpdatePendingChanges();
            }
        }
    }

    public string? AutoReplyStartTime
    {
        get => _autoReplyStartTime;
        set
        {
            if (SetProperty(ref _autoReplyStartTime, value))
            {
                UpdatePendingChanges();
            }
        }
    }

    public DateTime? AutoReplyEndDate
    {
        get => _autoReplyEndDate;
        set
        {
            if (SetProperty(ref _autoReplyEndDate, value))
            {
                UpdatePendingChanges();
            }
        }
    }

    public string? AutoReplyEndTime
    {
        get => _autoReplyEndTime;
        set
        {
            if (SetProperty(ref _autoReplyEndTime, value))
            {
                UpdatePendingChanges();
            }
        }
    }

    public string? AutoReplyInternalMessage
    {
        get => _autoReplyInternalMessage;
        set
        {
            if (SetProperty(ref _autoReplyInternalMessage, value))
            {
                UpdatePendingChanges();
            }
        }
    }

    public string? AutoReplyExternalMessage
    {
        get => _autoReplyExternalMessage;
        set
        {
            if (SetProperty(ref _autoReplyExternalMessage, value))
            {
                UpdatePendingChanges();
            }
        }
    }

    public string? AutoReplyExternalAudience
    {
        get => _autoReplyExternalAudience;
        set
        {
            if (SetProperty(ref _autoReplyExternalAudience, value))
            {
                UpdatePendingChanges();
            }
        }
    }

    public string AutoReplyStateLabel => AutoReplyEnabled
        ? AutoReplyScheduled ? "Programmato" : "Attivo"
        : "Disattivato";

    public void Initialize(AutoReplyConfigurationDto? autoReply)
    {
        _isInitializing = true;

        if (autoReply != null)
        {
            var state = autoReply.AutoReplyState ?? "Disabled";
            AutoReplyEnabled = !string.Equals(state, "Disabled", StringComparison.OrdinalIgnoreCase);
            AutoReplyScheduled = string.Equals(state, "Scheduled", StringComparison.OrdinalIgnoreCase);
            AutoReplyStartDate = autoReply.StartTime?.Date;
            AutoReplyStartTime = autoReply.StartTime?.ToString("HH:mm");
            AutoReplyEndDate = autoReply.EndTime?.Date;
            AutoReplyEndTime = autoReply.EndTime?.ToString("HH:mm");
            AutoReplyInternalMessage = NormalizeMessage(autoReply.InternalMessage);
            AutoReplyExternalMessage = NormalizeMessage(autoReply.ExternalMessage);
            AutoReplyExternalAudience = string.IsNullOrWhiteSpace(autoReply.ExternalAudience) ? "All" : autoReply.ExternalAudience;
        }
        else
        {
            AutoReplyEnabled = false;
            AutoReplyScheduled = false;
            AutoReplyStartDate = null;
            AutoReplyStartTime = null;
            AutoReplyEndDate = null;
            AutoReplyEndTime = null;
            AutoReplyInternalMessage = string.Empty;
            AutoReplyExternalMessage = string.Empty;
            AutoReplyExternalAudience = "All";
        }

        _originalSettings = CaptureSnapshot();
        _isInitializing = false;
        HasPendingChanges = false;
    }

    public void Reset()
    {
        _isInitializing = true;
        AutoReplyEnabled = false;
        AutoReplyScheduled = false;
        AutoReplyStartDate = null;
        AutoReplyStartTime = null;
        AutoReplyEndDate = null;
        AutoReplyEndTime = null;
        AutoReplyInternalMessage = string.Empty;
        AutoReplyExternalMessage = string.Empty;
        AutoReplyExternalAudience = "All";
        _originalSettings = null;
        _isInitializing = false;
        HasPendingChanges = false;
    }

    public void DiscardChanges()
    {
        if (_originalSettings == null)
        {
            return;
        }

        _isInitializing = true;
        AutoReplyEnabled = _originalSettings.AutoReplyEnabled;
        AutoReplyScheduled = _originalSettings.AutoReplyScheduled;
        AutoReplyStartDate = _originalSettings.AutoReplyStartDate?.Date;
        AutoReplyStartTime = _originalSettings.AutoReplyStartDate?.ToString("HH:mm");
        AutoReplyEndDate = _originalSettings.AutoReplyEndDate?.Date;
        AutoReplyEndTime = _originalSettings.AutoReplyEndDate?.ToString("HH:mm");
        AutoReplyInternalMessage = _originalSettings.AutoReplyInternalMessage;
        AutoReplyExternalMessage = _originalSettings.AutoReplyExternalMessage;
        AutoReplyExternalAudience = _originalSettings.AutoReplyExternalAudience;
        _isInitializing = false;
        HasPendingChanges = false;
    }

    public SetMailboxAutoReplyConfigurationRequest? BuildRequest(string identity, out string? validationError)
    {
        validationError = null;
        if (!HasPendingChanges)
        {
            return null;
        }

        if (AutoReplyScheduled)
        {
            if (!TryBuildScheduledDateTime(AutoReplyStartDate, AutoReplyStartTime, out var start) ||
                !TryBuildScheduledDateTime(AutoReplyEndDate, AutoReplyEndTime, out var end))
            {
                validationError = "Enter valid date and time values for the automatic reply window.";
                return null;
            }

            if (end <= start)
            {
                validationError = "La data of fine deve essere successiva alla data of inizio.";
                return null;
            }
        }

        return new SetMailboxAutoReplyConfigurationRequest
        {
            Identity = identity,
            AutoReplyState = AutoReplyEnabled
                ? AutoReplyScheduled ? "Scheduled" : "Enabled"
                : "Disabled",
            StartTime = AutoReplyScheduled ? BuildDateTime(AutoReplyStartDate, AutoReplyStartTime) : null,
            EndTime = AutoReplyScheduled ? BuildDateTime(AutoReplyEndDate, AutoReplyEndTime) : null,
            InternalMessage = NormalizeMessage(AutoReplyInternalMessage),
            ExternalMessage = NormalizeMessage(AutoReplyExternalMessage),
            ExternalAudience = AutoReplyEnabled ? AutoReplyExternalAudience : null
        };
    }

    private void UpdatePendingChanges()
    {
        if (_isInitializing || _originalSettings == null)
        {
            return;
        }

        HasPendingChanges = !CaptureSnapshot().Equals(_originalSettings);
    }

    private MailboxAutoReplySnapshot CaptureSnapshot()
    {
        return new MailboxAutoReplySnapshot
        {
            AutoReplyEnabled = AutoReplyEnabled,
            AutoReplyScheduled = AutoReplyScheduled,
            AutoReplyStartDate = BuildDateTime(AutoReplyStartDate, AutoReplyStartTime),
            AutoReplyEndDate = BuildDateTime(AutoReplyEndDate, AutoReplyEndTime),
            AutoReplyInternalMessage = NormalizeMessage(AutoReplyInternalMessage),
            AutoReplyExternalMessage = NormalizeMessage(AutoReplyExternalMessage),
            AutoReplyExternalAudience = NormalizeInput(AutoReplyExternalAudience)
        };
    }

    private static string NormalizeInput(string? value) => string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();

    private static string NormalizeMessage(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var normalized = Regex.Replace(value, @"<\s*br\s*/?>", "\n", RegexOptions.IgnoreCase);
        normalized = Regex.Replace(normalized, @"</\s*p\s*>", "\n", RegexOptions.IgnoreCase);
        normalized = Regex.Replace(normalized, @"</\s*div\s*>", "\n", RegexOptions.IgnoreCase);
        normalized = Regex.Replace(normalized, @"<\s*li\s*>", "- ", RegexOptions.IgnoreCase);
        normalized = Regex.Replace(normalized, @"</\s*li\s*>", "\n", RegexOptions.IgnoreCase);
        normalized = Regex.Replace(normalized, @"<[^>]+>", string.Empty);
        normalized = WebUtility.HtmlDecode(normalized);
        normalized = normalized.Replace("\r", string.Empty);
        normalized = Regex.Replace(normalized, @"\n{3,}", "\n\n");
        return normalized.Trim();
    }

    private static bool TryBuildScheduledDateTime(DateTime? date, string? timeText, out DateTime result)
    {
        result = default;
        if (date == null)
        {
            return false;
        }

        if (!TryParseTime(timeText, out var time))
        {
            return false;
        }

        result = date.Value.Date + time;
        return true;
    }

    private static DateTime? BuildDateTime(DateTime? date, string? timeText)
    {
        if (date == null)
        {
            return null;
        }

        if (!TryParseTime(timeText, out var time))
        {
            return date.Value.Date;
        }

        return date.Value.Date + time;
    }

    private static bool TryParseTime(string? timeText, out TimeSpan time)
    {
        time = default;
        if (string.IsNullOrWhiteSpace(timeText))
        {
            return false;
        }

        return TimeSpan.TryParseExact(
            timeText.Trim(),
            new[] { @"hh\:mm", @"h\:mm" },
            System.Globalization.CultureInfo.InvariantCulture,
            out time);
    }
}
