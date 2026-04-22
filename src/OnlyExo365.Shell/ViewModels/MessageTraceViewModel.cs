using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Windows.Input;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Shell.Helpers;
using Microsoft.Win32;

namespace OnlyExo365.Shell.ViewModels;

public class MessageTraceViewModel : ViewModelBase
{
    private const NavigationPage AlertPage = NavigationPage.MessageTrace;
    private readonly IWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;

    private CancellationTokenSource? _searchCts;

    private bool _isLoading;
    private string? _errorMessage;
    private string? _warningsText;
    private string? _diagnosticCorrelationId;
    private string? _senderAddress;
    private string? _recipientAddress;
    private DateTime _startDate = DateTime.Today.AddDays(-7);
    private DateTime _endDate = DateTime.Today;
    private int _currentPage = 1;
    private int _pageSize = 100;
    private bool _hasMore;
    private int _totalCount;
    private bool _isTotalCountExact;
    private double _loadingProgress;
    private string? _loadingStatus;
    private int? _loadingCurrentItem;
    private int? _loadingTotalItems;
    private string _statusFilter = "All";
    private MessageTraceItemDto? _selectedMessage;
    private bool _isLoadingDetails;
    private static readonly HashSet<string> AllowedStatusFilters = new(StringComparer.OrdinalIgnoreCase)
    {
        "All",
        "Delivered",
        "Failed",
        "Pending"
    };

    private readonly string _defaultExportDirectory = ExcelExportService.ResolveExportDirectory();

    public MessageTraceViewModel(IWorkerService workerService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;

        SearchCommand = new AsyncRelayCommand(SearchAsync, () => CanSearch);
        NextPageCommand = new AsyncRelayCommand(NextPageAsync, () => HasMore && !IsLoading);
        PreviousPageCommand = new AsyncRelayCommand(PreviousPageAsync, () => CurrentPage > 1 && !IsLoading);
        ExportExcelCommand = new RelayCommand(ExportExcel, () => Messages.Count > 0 && !IsLoading);
        LoadDetailsCommand = new AsyncRelayCommand(LoadDetailsAsync, () => SelectedMessage != null && !IsLoading && !IsLoadingDetails);
        SetStatusFilterCommand = new RelayCommand<string?>(SetStatusFilter);
    }

    public bool IsLoading
    {
        get => _isLoading;
        private set
        {
            if (SetProperty(ref _isLoading, value))
            {
                OnPropertyChanged(nameof(CanSearch));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool IsLoadingDetails
    {
        get => _isLoadingDetails;
        private set
        {
            if (SetProperty(ref _isLoadingDetails, value))
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

    public bool HasError => !string.IsNullOrEmpty(ErrorMessage);

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

    public string? SenderAddress
    {
        get => _senderAddress;
        set => SetProperty(ref _senderAddress, value);
    }

    public string? RecipientAddress
    {
        get => _recipientAddress;
        set => SetProperty(ref _recipientAddress, value);
    }

    public DateTime StartDate
    {
        get => _startDate;
        set => SetProperty(ref _startDate, value);
    }

    public DateTime EndDate
    {
        get => _endDate;
        set => SetProperty(ref _endDate, value);
    }

    public int CurrentPage
    {
        get => _currentPage;
        private set => SetProperty(ref _currentPage, value);
    }

    public int PageSize
    {
        get => _pageSize;
        set => SetProperty(ref _pageSize, value);
    }

    public bool HasMore
    {
        get => _hasMore;
        private set
        {
            if (SetProperty(ref _hasMore, value))
            {
                CommandManager.InvalidateRequerySuggested();
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
                OnPropertyChanged(nameof(ResultSummaryText));
            }
        }
    }

    public bool IsTotalCountExact
    {
        get => _isTotalCountExact;
        private set
        {
            if (SetProperty(ref _isTotalCountExact, value))
            {
                OnPropertyChanged(nameof(ResultSummaryText));
            }
        }
    }

    public double LoadingProgress
    {
        get => _loadingProgress;
        private set
        {
            if (SetProperty(ref _loadingProgress, value))
            {
                OnPropertyChanged(nameof(LoadingPercentText));
            }
        }
    }

    public string? LoadingStatus
    {
        get => _loadingStatus;
        private set => SetProperty(ref _loadingStatus, value);
    }

    public int? LoadingCurrentItem
    {
        get => _loadingCurrentItem;
        private set
        {
            if (SetProperty(ref _loadingCurrentItem, value))
            {
                OnPropertyChanged(nameof(HasLoadingCount));
                OnPropertyChanged(nameof(LoadingCountText));
            }
        }
    }

    public int? LoadingTotalItems
    {
        get => _loadingTotalItems;
        private set
        {
            if (SetProperty(ref _loadingTotalItems, value))
            {
                OnPropertyChanged(nameof(HasLoadingCount));
                OnPropertyChanged(nameof(LoadingCountText));
            }
        }
    }

    public bool HasLoadingCount => LoadingCurrentItem.HasValue;
    public string LoadingPercentText => FormatProgressPercent(LoadingProgress);
    public string? LoadingCountText => FormatProgressCount(LoadingCurrentItem, LoadingTotalItems, "messages");
    public string ResultSummaryText => IsTotalCountExact
        ? $"Total messages: {TotalCount}"
        : $"Messages found so far: {TotalCount}";

    public string StatusFilter
    {
        get => _statusFilter;
        set
        {
            if (SetProperty(ref _statusFilter, value))
            {
                ApplyStatusFilter();
            }
        }
    }

    public MessageTraceItemDto? SelectedMessage
    {
        get => _selectedMessage;
        set
        {
            if (SetProperty(ref _selectedMessage, value))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool CanSearch => !IsLoading && _shellViewModel.IsExchangeConnected;

    public ObservableCollection<MessageTraceItemDto> Messages { get; } = new();
    public ObservableCollection<MessageTraceItemDto> AllMessages { get; } = new();
    public ObservableCollection<MessageTraceDetailEventDto> SelectedMessageEvents { get; } = new();

    public ICommand SearchCommand { get; }
    public ICommand NextPageCommand { get; }
    public ICommand PreviousPageCommand { get; }
    public ICommand ExportExcelCommand { get; }
    public ICommand LoadDetailsCommand { get; }
    public ICommand SetStatusFilterCommand { get; }

    public Task LoadAsync()
    {
        OnPropertyChanged(nameof(CanSearch));
        CommandManager.InvalidateRequerySuggested();

        if (!_shellViewModel.IsExchangeConnected)
        {
            ErrorMessage = null;
            _shellViewModel.ClearPageAlert(AlertPage);
            return Task.CompletedTask;
        }

        ErrorMessage = null;
        _shellViewModel.ClearPageAlert(AlertPage);
        return Task.CompletedTask;
    }

    private async Task SearchAsync(CancellationToken cancellationToken)
    {
        CurrentPage = 1;
        await FetchMessagesAsync(cancellationToken);
    }

    private async Task NextPageAsync(CancellationToken cancellationToken)
    {
        CurrentPage++;
        await FetchMessagesAsync(cancellationToken);
    }

    private async Task PreviousPageAsync(CancellationToken cancellationToken)
    {
        if (CurrentPage > 1)
        {
            CurrentPage--;
            await FetchMessagesAsync(cancellationToken);
        }
    }

    private async Task FetchMessagesAsync(CancellationToken cancellationToken)
    {
        _searchCts?.Cancel();
        _searchCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

        IsLoading = true;
        ErrorMessage = null;
        WarningsText = null;
        DiagnosticCorrelationId = null;
        _shellViewModel.ClearPageAlert(AlertPage);
        LoadingProgress = 0;
        LoadingStatus = "Searching messages...";
        LoadingCurrentItem = null;
        LoadingTotalItems = null;
        IsTotalCountExact = false;

        try
        {
            var request = new GetMessageTraceRequest
            {
                SenderAddress = string.IsNullOrWhiteSpace(SenderAddress) ? null : SenderAddress.Trim(),
                RecipientAddress = string.IsNullOrWhiteSpace(RecipientAddress) ? null : RecipientAddress.Trim(),
                StartDate = StartDate,
                EndDate = EndDate.Date.AddDays(1).AddSeconds(-1),
                PageSize = PageSize,
                Page = CurrentPage
            };

            LoadingProgress = 30;
            LoadingStatus = "Querying Exchange...";

            var result = await _workerService.GetMessageTraceAsync(
                request,
                eventHandler: evt =>
                {
                    if (evt.EventType == EventType.Progress)
                    {
                        var progress = JsonMessageSerializer.ExtractPayload<ProgressEventPayload>(evt.Payload);
                        if (progress != null)
                        {
                            RunOnUiThread(() =>
                            {
                                LoadingProgress = progress.PercentComplete;
                                LoadingStatus = progress.StatusMessage;
                                LoadingCurrentItem = progress.CurrentItem;
                                LoadingTotalItems = progress.TotalItems;
                            });
                        }
                    }
                },
                cancellationToken: _searchCts.Token);

            LoadingProgress = 90;
            LoadingStatus = "Processing results...";

            if (result.IsSuccess && result.Value != null)
            {
                DiagnosticCorrelationId = result.Value.CorrelationId ?? result.CorrelationId;
                WarningsText = result.Value.Warnings.Count == 0
                    ? null
                    : string.Join(Environment.NewLine, result.Value.Warnings);

                RunOnUiThread(() =>
                {
                    AllMessages.Clear();
                    foreach (var msg in result.Value.Messages)
                    {
                        AllMessages.Add(msg);
                    }

                    ApplyStatusFilter();
                    TotalCount = result.Value.TotalCount;
                    IsTotalCountExact = result.Value.IsTotalCountExact;
                    HasMore = result.Value.HasMore;
                    SelectedMessage = null;
                    SelectedMessageEvents.Clear();
                });
                _shellViewModel.AddLog(LogLevel.Information, $"Message trace: {result.Value.Messages.Count} results found", correlationId: DiagnosticCorrelationId);
                foreach (var warning in result.Value.Warnings)
                {
                    _shellViewModel.AddLog(LogLevel.Warning, warning, correlationId: DiagnosticCorrelationId);
                }
            }
            else if (!result.WasCancelled)
            {
                DiagnosticCorrelationId = result.CorrelationId;
                var errorMessage = result.Error?.Message ?? "Unable to retrieve the message trace.";
                ErrorMessage = null;
                _shellViewModel.ShowPageLoadFailedAlert(AlertPage, errorMessage);
                _shellViewModel.AddLog(LogLevel.Error, $"Message trace failed: {errorMessage}", correlationId: DiagnosticCorrelationId);
            }

            LoadingProgress = 100;
            LoadingStatus = null;
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            ErrorMessage = null;
            _shellViewModel.ShowPageLoadFailedAlert(AlertPage, ex.Message);
            _shellViewModel.AddLog(LogLevel.Error, $"Message trace error: {ex.Message}", correlationId: DiagnosticCorrelationId);
        }
        finally
        {
            IsLoading = false;
            LoadingProgress = 0;
            LoadingStatus = null;
            LoadingCurrentItem = null;
            LoadingTotalItems = null;
        }
    }

    private async Task LoadDetailsAsync(CancellationToken cancellationToken)
    {
        if (SelectedMessage == null)
        {
            return;
        }

        IsLoadingDetails = true;
        ErrorMessage = null;

        try
        {
            var request = new GetMessageTraceDetailsRequest
            {
                MessageTraceId = SelectedMessage.MessageTraceId,
                RecipientAddress = SelectedMessage.RecipientAddress
            };

            var result = await _workerService.GetMessageTraceDetailsAsync(request, cancellationToken: cancellationToken);
            if (result.IsSuccess && result.Value != null)
            {
                DiagnosticCorrelationId = result.Value.CorrelationId ?? result.CorrelationId ?? DiagnosticCorrelationId;
                if (result.Value.Warnings.Count > 0)
                {
                    var warningBlock = string.Join(Environment.NewLine, result.Value.Warnings);
                    WarningsText = string.IsNullOrWhiteSpace(WarningsText)
                        ? warningBlock
                        : WarningsText + Environment.NewLine + warningBlock;

                    foreach (var warning in result.Value.Warnings)
                    {
                        _shellViewModel.AddLog(LogLevel.Warning, warning, correlationId: DiagnosticCorrelationId);
                    }
                }

                SelectedMessageEvents.Clear();
                foreach (var item in result.Value.Events.OrderBy(e => e.Date))
                {
                    SelectedMessageEvents.Add(item);
                }
            }
            else if (!result.WasCancelled)
            {
                DiagnosticCorrelationId = result.CorrelationId ?? DiagnosticCorrelationId;
                ErrorMessage = result.Error?.Message ?? "Unable to retrieve message details.";
            }
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsLoadingDetails = false;
        }
    }

    private void SetStatusFilter(string? status)
    {
        var normalized = string.IsNullOrWhiteSpace(status) ? "All" : status.Trim();
        StatusFilter = AllowedStatusFilters.Contains(normalized) ? normalized : "All";
    }

    private void ApplyStatusFilter()
    {
        var normalizedFilter = string.IsNullOrWhiteSpace(StatusFilter) ? "All" : StatusFilter.Trim();
        if (!AllowedStatusFilters.Contains(normalizedFilter))
        {
            normalizedFilter = "All";
            if (!string.Equals(StatusFilter, normalizedFilter, StringComparison.Ordinal))
            {
                _statusFilter = normalizedFilter;
                OnPropertyChanged(nameof(StatusFilter));
            }
        }

        IEnumerable<MessageTraceItemDto> filtered = AllMessages;
        if (!string.Equals(normalizedFilter, "All", StringComparison.OrdinalIgnoreCase))
        {
            filtered = filtered.Where(m => string.Equals(m.Status?.Trim(), normalizedFilter, StringComparison.OrdinalIgnoreCase));
        }

        Messages.Clear();
        foreach (var item in filtered)
        {
            Messages.Add(item);
        }

        CommandManager.InvalidateRequerySuggested();
    }

    private void ExportExcel()
    {
        if (Messages.Count == 0)
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
                FileName = $"message-trace-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx"
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            ExcelExportService.ExportWorkbook(
                dialog.FileName,
                "MessageTrace",
                ["Received", "Sender", "Recipient", "Subject", "Status", "Size", "MessageId", "MessageTraceId"],
                Messages.Select(message => (IReadOnlyList<string?>)
                [
                    message.Received?.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
                    message.SenderAddress,
                    message.RecipientAddress,
                    message.Subject,
                    message.Status,
                    message.Size?.ToString(CultureInfo.InvariantCulture),
                    message.MessageId,
                    message.MessageTraceId
                ]));

            _shellViewModel.AddLog(LogLevel.Information, $"Message trace export Excel saved: {dialog.FileName}");
        }
        catch (Exception ex)
        {
            ErrorMessage = $"Export failed: {ex.Message}";
            _shellViewModel.AddLog(LogLevel.Error, $"Message trace export failed: {ex.Message}");
        }
    }

    public void Cancel()
    {
        _searchCts?.Cancel();
    }
}

