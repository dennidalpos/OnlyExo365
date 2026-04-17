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

public class MailboxSpaceViewModel : ViewModelBase
{
    private const NavigationPage AlertPage = NavigationPage.MailboxSpace;
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

    public MailboxSpaceViewModel(IMailboxesWorkerService workerService, NavigationService navigationService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;

        StartScanCommand = new AsyncRelayCommand(StartScanAsync, () => !IsLoading && _shellViewModel.IsExchangeConnected);
        ExportExcelCommand = new RelayCommand(ExportExcel, () => Mailboxes.Count > 0 && !IsLoading);
        _shellViewModel.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName == nameof(ShellViewModel.IsExchangeConnected))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        };
    }

    public ObservableCollection<MailboxSpaceItemViewModel> Mailboxes { get; } = new();

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

    private async Task StartScanAsync(CancellationToken cancellationToken)
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            ErrorMessage = null;
            _shellViewModel.ClearPageAlert(AlertPage);
            return;
        }

        IsLoading = true;
        ErrorMessage = null;
        WarningsText = null;
        DiagnosticCorrelationId = null;
        Mailboxes.Clear();
        _shellViewModel.ClearPageAlert(AlertPage);
        ProgressPercent = 0;
        ProgressStatus = "Starting scan...";
        ProgressCurrentItem = null;
        ProgressTotalItems = null;

        try
        {
            var result = await _workerService.GetMailboxSpaceReportAsync(
                new GetMailboxSpaceReportRequest(),
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

                var items = result.Value.Mailboxes
                    .Select(item => new MailboxSpaceItemViewModel(item))
                    .OrderBy(item => item.RemainingPercent ?? double.MaxValue)
                    .ToList();

                foreach (var item in items)
                {
                    Mailboxes.Add(item);
                }

                foreach (var warning in result.Value.Warnings)
                {
                    _shellViewModel.AddLog(LogLevel.Warning, warning, correlationId: DiagnosticCorrelationId);
                }
            }
            else if (!result.WasCancelled)
            {
                DiagnosticCorrelationId = result.CorrelationId;
                var errorMessage = result.Error?.Message ?? "Unable to load the mailbox storage report.";
                ErrorMessage = null;
                _shellViewModel.ShowPageLoadFailedAlert(AlertPage, errorMessage);
                _shellViewModel.AddLog(LogLevel.Error, errorMessage, correlationId: DiagnosticCorrelationId);
            }
        }
        catch (Exception ex)
        {
            ErrorMessage = null;
            _shellViewModel.ShowPageLoadFailedAlert(AlertPage, ex.Message);
            _shellViewModel.AddLog(LogLevel.Error, $"Mailbox space scan error: {ex.Message}", correlationId: DiagnosticCorrelationId);
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

    private void ExportExcel()
    {
        if (Mailboxes.Count == 0)
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
                FileName = $"mailbox-space-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx"
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            ExcelExportService.ExportWorkbook(
                dialog.FileName,
                "MailboxSpace",
                ["Identity", "DisplayName", "PrimarySmtpAddress", "Used", "Quota", "RemainingPercent", "RemainingCategory"],
                Mailboxes.Select(mailbox => (IReadOnlyList<string?>)
                [
                    mailbox.Identity,
                    mailbox.DisplayName,
                    mailbox.PrimarySmtpAddress,
                    mailbox.TotalItemSize,
                    mailbox.QuotaLabel,
                    mailbox.RemainingPercentDisplay,
                    mailbox.RemainingCategory
                ]));

            _shellViewModel.AddLog(LogLevel.Information, $"Mailbox space export Excel saved: {dialog.FileName}");
        }
        catch (Exception ex)
        {
            ErrorMessage = $"Export failed: {ex.Message}";
            _shellViewModel.AddLog(LogLevel.Error, $"Mailbox space export failed: {ex.Message}", correlationId: DiagnosticCorrelationId);
        }
    }
}

public class MailboxSpaceItemViewModel
{
    public MailboxSpaceItemViewModel(MailboxSpaceItemDto dto)
    {
        Identity = dto.Identity;
        DisplayName = dto.DisplayName;
        PrimarySmtpAddress = dto.PrimarySmtpAddress;
        TotalItemSize = dto.TotalItemSize ?? "-";
        QuotaLabel = dto.ProhibitSendReceiveQuota ?? dto.ProhibitSendQuota ?? dto.IssueWarningQuota ?? "-";

        var quotaBytes = dto.ProhibitSendReceiveQuotaBytes
            ?? dto.ProhibitSendQuotaBytes
            ?? dto.IssueWarningQuotaBytes;

        if (quotaBytes.HasValue && quotaBytes.Value > 0 && dto.TotalItemSizeBytes.HasValue)
        {
            RemainingPercent = (quotaBytes.Value - dto.TotalItemSizeBytes.Value) / (double)quotaBytes.Value * 100.0;
        }
        else
        {
            RemainingPercent = null;
        }

        RemainingCategory = RemainingPercent switch
        {
            null => "Unknown",
            < 5 => "Critical",
            <= 15 => "Warning",
            _ => "Ok"
        };
    }

    public string Identity { get; }
    public string DisplayName { get; }
    public string PrimarySmtpAddress { get; }
    public string TotalItemSize { get; }
    public string QuotaLabel { get; }
    public double? RemainingPercent { get; }
    public string RemainingCategory { get; }
    public string RemainingPercentDisplay => RemainingPercent.HasValue ? $"{RemainingPercent.Value:0.0}%" : "N/D";
}

