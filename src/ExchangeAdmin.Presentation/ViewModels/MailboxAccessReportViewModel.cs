using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;
using Microsoft.Win32;

namespace ExchangeAdmin.Presentation.ViewModels;

public class MailboxAccessReportViewModel : ViewModelBase
{
    private readonly IMailboxesWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;
    private bool _isLoading;
    private string? _errorMessage;
    private string? _warningsText;
    private string? _diagnosticCorrelationId;
    private double _progressPercent;
    private string? _progressStatus;
    private int? _progressCurrentItem;
    private int? _progressTotalItems;
    private readonly string _defaultExportDirectory = ExcelExportService.ResolveExportDirectory();

    public MailboxAccessReportViewModel(IMailboxesWorkerService workerService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;

        StartScanCommand = new AsyncRelayCommand(LoadAsync, () => !IsLoading && _shellViewModel.IsExchangeConnected);
        ExportExcelCommand = new RelayCommand(ExportExcel, () => Rows.Count > 0 && !IsLoading);
        _shellViewModel.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName == nameof(ShellViewModel.IsExchangeConnected))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        };
    }

    public ObservableCollection<MailboxAccessMatrixRowViewModel> Rows { get; } = new();

    public bool IsLoading
    {
        get => _isLoading;
        private set
        {
            if (SetProperty(ref _isLoading, value))
            {
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

    public string? WarningsText
    {
        get => _warningsText;
        private set
        {
            if (SetProperty(ref _warningsText, value))
            {
                OnPropertyChanged(nameof(HasWarnings));
            }
        }
    }

    public bool HasWarnings => !string.IsNullOrWhiteSpace(WarningsText);

    public string? DiagnosticCorrelationId
    {
        get => _diagnosticCorrelationId;
        private set
        {
            if (SetProperty(ref _diagnosticCorrelationId, value))
            {
                OnPropertyChanged(nameof(HasDiagnosticReference));
                OnPropertyChanged(nameof(DiagnosticReferenceText));
            }
        }
    }

    public bool HasDiagnosticReference => !string.IsNullOrWhiteSpace(DiagnosticCorrelationId);
    public string? DiagnosticReferenceText => HasDiagnosticReference ? $"Ref: {DiagnosticCorrelationId}" : null;

    public double ProgressPercent
    {
        get => _progressPercent;
        private set
        {
            if (SetProperty(ref _progressPercent, value))
            {
                OnPropertyChanged(nameof(ProgressPercentText));
            }
        }
    }

    public string? ProgressStatus
    {
        get => _progressStatus;
        private set => SetProperty(ref _progressStatus, value);
    }

    public int? ProgressCurrentItem
    {
        get => _progressCurrentItem;
        private set
        {
            if (SetProperty(ref _progressCurrentItem, value))
            {
                OnPropertyChanged(nameof(HasProgressCount));
                OnPropertyChanged(nameof(ProgressCountText));
            }
        }
    }

    public int? ProgressTotalItems
    {
        get => _progressTotalItems;
        private set
        {
            if (SetProperty(ref _progressTotalItems, value))
            {
                OnPropertyChanged(nameof(HasProgressCount));
                OnPropertyChanged(nameof(ProgressCountText));
            }
        }
    }

    public bool HasProgressCount => ProgressCurrentItem.HasValue;
    public string ProgressPercentText => FormatProgressPercent(ProgressPercent);
    public string? ProgressCountText => FormatProgressCount(ProgressCurrentItem, ProgressTotalItems, "mailboxes");

    public ICommand StartScanCommand { get; }
    public ICommand ExportExcelCommand { get; }

    public async Task LoadAsync(CancellationToken cancellationToken = default)
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            ErrorMessage = "Not connected to Exchange Online";
            return;
        }

        IsLoading = true;
        ErrorMessage = null;
        WarningsText = null;
        DiagnosticCorrelationId = null;
        Rows.Clear();
        ProgressPercent = 0;
        ProgressStatus = "Starting mailbox access scan...";
        ProgressCurrentItem = null;
        ProgressTotalItems = null;

        try
        {
            var result = await _workerService.GetMailboxAccessReportAsync(
                new GetMailboxAccessReportRequest(),
                eventHandler: evt =>
                {
                    if (evt.EventType != EventType.Progress)
                    {
                        return;
                    }

                    var progress = JsonMessageSerializer.ExtractPayload<ProgressEventPayload>(evt.Payload);
                    if (progress == null)
                    {
                        return;
                    }

                    RunOnUiThread(() =>
                    {
                        ProgressPercent = progress.PercentComplete;
                        ProgressStatus = progress.StatusMessage;
                        ProgressCurrentItem = progress.CurrentItem;
                        ProgressTotalItems = progress.TotalItems;
                    });
                },
                cancellationToken: cancellationToken);

            if (result.IsSuccess && result.Value != null)
            {
                DiagnosticCorrelationId = result.Value.CorrelationId ?? result.CorrelationId;
                WarningsText = result.Value.Warnings.Count == 0
                    ? null
                    : string.Join(Environment.NewLine, result.Value.Warnings);

                var rows = result.Value.Grants
                    .Where(grant => !string.IsNullOrWhiteSpace(grant.User))
                    .GroupBy(grant => new
                    {
                        User = grant.User.Trim(),
                        Mailbox = SelectMailboxLabel(grant)
                    })
                    .Select(group => new MailboxAccessMatrixRowViewModel(group.Key.User, group.Key.Mailbox, group.ToList()))
                    .OrderBy(row => row.User)
                    .ThenBy(row => row.Mailbox)
                    .ToList();

                foreach (var row in rows)
                {
                    Rows.Add(row);
                }

                foreach (var warning in result.Value.Warnings)
                {
                    _shellViewModel.AddLog(LogLevel.Warning, warning, correlationId: DiagnosticCorrelationId);
                }
            }
            else if (!result.WasCancelled)
            {
                DiagnosticCorrelationId = result.CorrelationId;
                ErrorMessage = result.Error?.Message ?? "Unable to load the mailbox access report.";
                _shellViewModel.AddLog(LogLevel.Error, ErrorMessage, correlationId: DiagnosticCorrelationId);
            }
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Mailbox access report error: {ex.Message}", correlationId: DiagnosticCorrelationId);
        }
        finally
        {
            IsLoading = false;
            ProgressStatus = null;
            ProgressPercent = 0;
            ProgressCurrentItem = null;
            ProgressTotalItems = null;
        }
    }

    private static string SelectMailboxLabel(MailboxAccessGrantDto grant)
    {
        if (!string.IsNullOrWhiteSpace(grant.MailboxPrimarySmtpAddress))
        {
            return grant.MailboxPrimarySmtpAddress;
        }

        if (!string.IsNullOrWhiteSpace(grant.MailboxDisplayName))
        {
            return grant.MailboxDisplayName;
        }

        return grant.MailboxIdentity;
    }

    private void ExportExcel()
    {
        if (Rows.Count == 0)
        {
            return;
        }

        try
        {
            Directory.CreateDirectory(_defaultExportDirectory);

            var dialog = new SaveFileDialog
            {
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                InitialDirectory = _defaultExportDirectory,
                FileName = $"mailbox-access-report-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx"
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            ExcelExportService.ExportWorkbook(
                dialog.FileName,
                "MailboxAccess",
                ["User", "Mailbox", "FullAccess", "SendAs", "SendOnBehalf"],
                Rows.Select(row => (IReadOnlyList<string?>)
                [
                    row.User,
                    row.Mailbox,
                    row.FullAccess,
                    row.SendAs,
                    row.SendOnBehalf
                ]));

            _shellViewModel.AddLog(LogLevel.Information, $"Mailbox access report export Excel saved: {dialog.FileName}");
        }
        catch (Exception ex)
        {
            ErrorMessage = $"Export failed: {ex.Message}";
            _shellViewModel.AddLog(LogLevel.Error, $"Mailbox access report export failed: {ex.Message}", correlationId: DiagnosticCorrelationId);
        }
    }
}

public class MailboxAccessMatrixRowViewModel
{
    public MailboxAccessMatrixRowViewModel(string user, string mailbox, IReadOnlyCollection<MailboxAccessGrantDto> grants)
    {
        User = user;
        Mailbox = mailbox;
        FullAccess = BuildPermissionCell(grants, "FullAccess");
        SendAs = BuildPermissionCell(grants, "SendAs");
        SendOnBehalf = BuildPermissionCell(grants, "SendOnBehalf");
    }

    public string User { get; }

    public string Mailbox { get; }

    public string FullAccess { get; }

    public string SendAs { get; }

    public string SendOnBehalf { get; }

    private static string BuildPermissionCell(IEnumerable<MailboxAccessGrantDto> grants, string permissionType)
    {
        var matching = grants
            .Where(grant => string.Equals(grant.PermissionType, permissionType, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (matching.Count == 0)
        {
            return "-";
        }

        var normalizedRights = matching
            .SelectMany(grant => grant.AccessRights)
            .Where(right => !string.IsNullOrWhiteSpace(right))
            .Select(right => right.Trim())
            .Where(right => !string.Equals(right, permissionType, StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(right => right)
            .ToList();

        return normalizedRights.Count == 0
            ? "Yes"
            : string.Join(", ", normalizedRights);
    }
}

