using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

public class MailboxSpaceViewModel : ViewModelBase
{
    private readonly IWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;
    private bool _isLoading;
    private string? _errorMessage;
    private double _progressPercent;
    private string? _progressStatus;

    public MailboxSpaceViewModel(IWorkerService workerService, NavigationService navigationService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;

        StartScanCommand = new AsyncRelayCommand(StartScanAsync, () => !IsLoading && _shellViewModel.IsExchangeConnected);
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

    public double ProgressPercent
    {
        get => _progressPercent;
        private set => SetProperty(ref _progressPercent, value);
    }

    public string? ProgressStatus
    {
        get => _progressStatus;
        private set => SetProperty(ref _progressStatus, value);
    }

    public ICommand StartScanCommand { get; }

    private async Task StartScanAsync(CancellationToken cancellationToken)
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            ErrorMessage = "Non connesso a Exchange Online";
            return;
        }

        IsLoading = true;
        ErrorMessage = null;
        Mailboxes.Clear();
        ProgressPercent = 0;
        ProgressStatus = "Avvio scansione...";

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
                    });
                },
                cancellationToken: cancellationToken);

            if (result.IsSuccess && result.Value != null)
            {
                var items = result.Value.Mailboxes
                    .Select(item => new MailboxSpaceItemViewModel(item))
                    .OrderBy(item => item.RemainingPercent ?? double.MaxValue)
                    .ToList();

                foreach (var item in items)
                {
                    Mailboxes.Add(item);
                }
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? "Impossibile caricare il report spazio mailbox";
                _shellViewModel.AddLog(LogLevel.Error, ErrorMessage);
            }
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Mailbox space scan error: {ex.Message}");
        }
        finally
        {
            IsLoading = false;
            ProgressStatus = null;
            ProgressPercent = 0;
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
