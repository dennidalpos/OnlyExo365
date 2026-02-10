using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows.Input;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;
using Microsoft.Win32;

namespace ExchangeAdmin.Presentation.ViewModels;

public class MessageTraceViewModel : ViewModelBase
{
    private readonly IWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;

    private CancellationTokenSource? _searchCts;

    private bool _isLoading;
    private string? _errorMessage;
    private string? _senderAddress;
    private string? _recipientAddress;
    private DateTime _startDate = DateTime.Today.AddDays(-7);
    private DateTime _endDate = DateTime.Today;
    private int _currentPage = 1;
    private int _pageSize = 100;
    private bool _hasMore;
    private int _totalCount;
    private double _loadingProgress;
    private string? _loadingStatus;
    private string _statusFilter = "All";
    private MessageTraceItemDto? _selectedMessage;
    private bool _isLoadingDetails;

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
        private set => SetProperty(ref _totalCount, value);
    }

    public double LoadingProgress
    {
        get => _loadingProgress;
        private set => SetProperty(ref _loadingProgress, value);
    }

    public string? LoadingStatus
    {
        get => _loadingStatus;
        private set => SetProperty(ref _loadingStatus, value);
    }

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

    public async Task LoadAsync()
    {
        OnPropertyChanged(nameof(CanSearch));
        CommandManager.InvalidateRequerySuggested();

        if (!_shellViewModel.IsExchangeConnected)
        {
            ErrorMessage = "Non connesso a Exchange Online";
            return;
        }

        ErrorMessage = null;
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
        LoadingProgress = 0;
        LoadingStatus = "Ricerca messaggi...";

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
            LoadingStatus = "Interrogazione Exchange...";

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
                            });
                        }
                    }
                },
                cancellationToken: _searchCts.Token);

            LoadingProgress = 90;
            LoadingStatus = "Elaborazione risultati...";

            if (result.IsSuccess && result.Value != null)
            {
                RunOnUiThread(() =>
                {
                    AllMessages.Clear();
                    foreach (var msg in result.Value.Messages)
                    {
                        AllMessages.Add(msg);
                    }

                    ApplyStatusFilter();
                    TotalCount = result.Value.TotalCount;
                    HasMore = result.Value.HasMore;
                    SelectedMessage = null;
                    SelectedMessageEvents.Clear();
                });
                _shellViewModel.AddLog(LogLevel.Information, $"Message trace: {result.Value.Messages.Count} risultati trovati");
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? "Impossibile recuperare la traccia messaggi";
                _shellViewModel.AddLog(LogLevel.Error, $"Message trace failed: {ErrorMessage}");
            }

            LoadingProgress = 100;
            LoadingStatus = null;
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Message trace error: {ex.Message}");
        }
        finally
        {
            IsLoading = false;
            LoadingProgress = 0;
            LoadingStatus = null;
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
                SelectedMessageEvents.Clear();
                foreach (var item in result.Value.Events.OrderBy(e => e.Date))
                {
                    SelectedMessageEvents.Add(item);
                }
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? "Impossibile recuperare il dettaglio messaggio";
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
        StatusFilter = string.IsNullOrWhiteSpace(status) ? "All" : status;
    }

    private void ApplyStatusFilter()
    {
        IEnumerable<MessageTraceItemDto> filtered = AllMessages;
        if (!string.Equals(StatusFilter, "All", StringComparison.OrdinalIgnoreCase))
        {
            filtered = filtered.Where(m => string.Equals(m.Status, StatusFilter, StringComparison.OrdinalIgnoreCase));
        }

        Messages.Clear();
        foreach (var item in filtered)
        {
            Messages.Add(item);
        }
    }

    private void ExportExcel()
    {
        if (Messages.Count == 0)
        {
            return;
        }

        var dialog = new SaveFileDialog
        {
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            FileName = $"message-trace-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx"
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        using var spreadsheet = SpreadsheetDocument.Create(dialog.FileName, SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(new SheetViews(new SheetView { WorkbookViewId = 0U }), sheetData);

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = CreateStylesheet();
        stylesPart.Stylesheet.Save();

        var sheets = spreadsheet.WorkbookPart?.Workbook.AppendChild(new Sheets());
        var sheet = new Sheet
        {
            Id = spreadsheet.WorkbookPart?.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "MessageTrace"
        };
        sheets?.Append(sheet);

        var header = new Row();
        foreach (var title in new[] { "Received", "Sender", "Recipient", "Subject", "Status", "Size", "MessageId", "MessageTraceId" })
        {
            header.Append(CreateTextCell(title, 1U));
        }
        sheetData.Append(header);

        foreach (var message in Messages)
        {
            var row = new Row();
            row.Append(CreateTextCell(message.Received?.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) ?? string.Empty));
            row.Append(CreateTextCell(message.SenderAddress));
            row.Append(CreateTextCell(message.RecipientAddress));
            row.Append(CreateTextCell(message.Subject));
            row.Append(CreateTextCell(message.Status));
            row.Append(CreateTextCell(message.Size?.ToString(CultureInfo.InvariantCulture) ?? string.Empty));
            row.Append(CreateTextCell(message.MessageId));
            row.Append(CreateTextCell(message.MessageTraceId));
            sheetData.Append(row);
        }

        worksheetPart.Worksheet.Save();
        workbookPart.Workbook.Save();

        _shellViewModel.AddLog(LogLevel.Information, $"Message trace export Excel salvato: {dialog.FileName}");
    }

    private static Stylesheet CreateStylesheet()
    {
        var fonts = new Fonts(
            new Font(),
            new Font(new Bold()));

        var fills = new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
            new Fill(new PatternFill(new ForegroundColor { Rgb = HexBinaryValue.FromString("FFDDEBF7") }) { PatternType = PatternValues.Solid }));

        var borders = new Borders(new Border());

        var cellFormats = new CellFormats(
            new CellFormat(),
            new CellFormat
            {
                FontId = 1,
                FillId = 2,
                BorderId = 0,
                ApplyFont = true,
                ApplyFill = true
            });

        return new Stylesheet(fonts, fills, borders, cellFormats);
    }

    private static Cell CreateTextCell(string? value, uint styleIndex = 0)
    {
        return new Cell
        {
            DataType = CellValues.InlineString,
            StyleIndex = styleIndex,
            InlineString = new InlineString(new Text(value ?? string.Empty))
        };
    }

    public void Cancel()
    {
        _searchCts?.Cancel();
    }
}
