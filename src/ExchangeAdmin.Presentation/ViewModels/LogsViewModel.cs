using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Windows.Input;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;

namespace ExchangeAdmin.Presentation.ViewModels;

/// <summary>
/// ViewModel per il visualizzatore log con buffering e auto-scroll intelligente.
/// </summary>
public sealed class LogsViewModel : ViewModelBase, IDisposable
{
    private readonly ShellViewModel _shellViewModel;
    private readonly object _filterLock = new();

    private LogLevel _filterLevel = LogLevel.Verbose;
    private string? _searchFilter;
    private bool _autoScroll = true;
    private bool _userHasScrolled;
    private bool _isRefreshing;
    private int _pendingUpdateCount;

    // Buffered filtered logs per ridurre allocazioni
    private List<LogEntry>? _cachedFilteredLogs;
    private bool _filterCacheInvalid = true;

    // Throttling per evitare troppi refresh
    private DateTime _lastRefreshTime = DateTime.MinValue;
    private static readonly TimeSpan MinRefreshInterval = TimeSpan.FromMilliseconds(100);

    /// <summary>
    /// Crea un nuovo ViewModel per i log.
    /// </summary>
    /// <param name="shellViewModel">Shell ViewModel che contiene i log.</param>
    public LogsViewModel(ShellViewModel shellViewModel)
    {
        _shellViewModel = shellViewModel;

        ClearLogsCommand = new RelayCommand(() =>
        {
            _shellViewModel.LogEntries.Clear();
            InvalidateFilterCache();
        });

        CopyLogsCommand = new RelayCommand(CopyLogs);
        ScrollToTopCommand = new RelayCommand(ScrollToTop);
        ScrollToBottomCommand = new RelayCommand(ScrollToBottom);

        // Subscribe ai cambiamenti della collection
        _shellViewModel.LogEntries.CollectionChanged += OnLogEntriesChanged;
    }

    #region Properties

    /// <summary>
    /// Collection originale dei log dal ShellViewModel.
    /// </summary>
    public ObservableCollection<LogEntry> LogEntries => _shellViewModel.LogEntries;

    /// <summary>
    /// Livello minimo di log da visualizzare.
    /// </summary>
    public LogLevel FilterLevel
    {
        get => _filterLevel;
        set
        {
            if (SetProperty(ref _filterLevel, value))
            {
                InvalidateFilterCache();
                NotifyFilteredLogsChanged();
            }
        }
    }

    /// <summary>
    /// Filtro di ricerca testuale.
    /// </summary>
    public string? SearchFilter
    {
        get => _searchFilter;
        set
        {
            if (SetProperty(ref _searchFilter, value))
            {
                InvalidateFilterCache();
                NotifyFilteredLogsChanged();
            }
        }
    }

    /// <summary>
    /// Se true, auto-scroll verso nuovi log (solo se utente non ha scrollato manualmente).
    /// </summary>
    public bool AutoScroll
    {
        get => _autoScroll;
        set
        {
            if (SetProperty(ref _autoScroll, value))
            {
                if (value)
                {
                    _userHasScrolled = false;
                }
            }
        }
    }

    /// <summary>
    /// Indica se l'utente ha scrollato manualmente (disabilita auto-scroll temporaneamente).
    /// </summary>
    public bool UserHasScrolled
    {
        get => _userHasScrolled;
        set => SetProperty(ref _userHasScrolled, value);
    }

    /// <summary>
    /// Indica se è necessario fare auto-scroll (auto-scroll abilitato e utente non ha scrollato).
    /// </summary>
    public bool ShouldAutoScroll => AutoScroll && !_userHasScrolled;

    /// <summary>
    /// Log filtrati con caching.
    /// </summary>
    public IEnumerable<LogEntry> FilteredLogs
    {
        get
        {
            lock (_filterLock)
            {
                if (_filterCacheInvalid || _cachedFilteredLogs == null)
                {
                    RebuildFilterCache();
                }
                return _cachedFilteredLogs ?? Enumerable.Empty<LogEntry>();
            }
        }
    }

    /// <summary>
    /// Numero totale di log entry.
    /// </summary>
    public int TotalCount => LogEntries.Count;

    /// <summary>
    /// Numero di log filtrati visibili.
    /// </summary>
    public int FilteredCount
    {
        get
        {
            lock (_filterLock)
            {
                return _cachedFilteredLogs?.Count ?? 0;
            }
        }
    }

    /// <summary>Conteggio log livello Verbose.</summary>
    public int VerboseCount => CountByLevel(LogLevel.Verbose);

    /// <summary>Conteggio log livello Debug.</summary>
    public int DebugCount => CountByLevel(LogLevel.Debug);

    /// <summary>Conteggio log livello Information.</summary>
    public int InfoCount => CountByLevel(LogLevel.Information);

    /// <summary>Conteggio log livello Warning.</summary>
    public int WarningCount => CountByLevel(LogLevel.Warning);

    /// <summary>Conteggio log livello Error.</summary>
    public int ErrorCount => CountByLevel(LogLevel.Error);

    /// <summary>
    /// Indica se ci sono log in errore.
    /// </summary>
    public bool HasErrors => ErrorCount > 0;

    /// <summary>
    /// Indica se ci sono warning.
    /// </summary>
    public bool HasWarnings => WarningCount > 0;

    #endregion

    #region Commands

    /// <summary>Cancella tutti i log.</summary>
    public ICommand ClearLogsCommand { get; }

    /// <summary>Copia i log filtrati negli appunti.</summary>
    public ICommand CopyLogsCommand { get; }

    /// <summary>Scrolla all'inizio dei log.</summary>
    public ICommand ScrollToTopCommand { get; }

    /// <summary>Scrolla alla fine dei log.</summary>
    public ICommand ScrollToBottomCommand { get; }

    #endregion

    #region Methods

    /// <summary>
    /// Forza un refresh dei log filtrati.
    /// </summary>
    public void Refresh()
    {
        InvalidateFilterCache();
        NotifyFilteredLogsChanged();
        NotifyCountsChanged();
    }

    /// <summary>
    /// Notifica che l'utente ha scrollato manualmente.
    /// </summary>
    public void NotifyUserScrolled()
    {
        if (AutoScroll)
        {
            _userHasScrolled = true;
            OnPropertyChanged(nameof(UserHasScrolled));
            OnPropertyChanged(nameof(ShouldAutoScroll));
        }
    }

    /// <summary>
    /// Resetta lo stato di scroll dell'utente (auto-scroll riprende).
    /// </summary>
    public void ResetUserScroll()
    {
        _userHasScrolled = false;
        OnPropertyChanged(nameof(UserHasScrolled));
        OnPropertyChanged(nameof(ShouldAutoScroll));
    }

    private void OnLogEntriesChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        _pendingUpdateCount++;
        InvalidateFilterCache();

        // Throttle updates per evitare troppi refresh durante bulk operations
        var now = DateTime.UtcNow;
        if (now - _lastRefreshTime >= MinRefreshInterval)
        {
            _lastRefreshTime = now;
            _pendingUpdateCount = 0;

            RunOnUiThread(() =>
            {
                NotifyFilteredLogsChanged();
                NotifyCountsChanged();
            });
        }
    }

    private void InvalidateFilterCache()
    {
        lock (_filterLock)
        {
            _filterCacheInvalid = true;
        }
    }

    private void RebuildFilterCache()
    {
        // Già dentro lock
        _cachedFilteredLogs = new List<LogEntry>();
        var searchLower = _searchFilter?.ToLowerInvariant();

        foreach (var log in LogEntries)
        {
            // Filtro per livello
            if (log.Level < _filterLevel)
            {
                continue;
            }

            // Filtro per ricerca
            if (!string.IsNullOrEmpty(searchLower))
            {
                var messageContains = log.Message.ToLowerInvariant().Contains(searchLower);
                var sourceContains = log.Source?.ToLowerInvariant().Contains(searchLower) ?? false;

                if (!messageContains && !sourceContains)
                {
                    continue;
                }
            }

            _cachedFilteredLogs.Add(log);
        }

        _filterCacheInvalid = false;
    }

    private int CountByLevel(LogLevel level)
    {
        var count = 0;
        foreach (var log in LogEntries)
        {
            if (log.Level == level)
            {
                count++;
            }
        }
        return count;
    }

    private void NotifyFilteredLogsChanged()
    {
        if (_isRefreshing)
        {
            return;
        }

        _isRefreshing = true;
        try
        {
            OnPropertyChanged(nameof(FilteredLogs));
            OnPropertyChanged(nameof(FilteredCount));
        }
        finally
        {
            _isRefreshing = false;
        }
    }

    private void NotifyCountsChanged()
    {
        OnPropertyChanged(nameof(TotalCount));
        OnPropertyChanged(nameof(VerboseCount));
        OnPropertyChanged(nameof(DebugCount));
        OnPropertyChanged(nameof(InfoCount));
        OnPropertyChanged(nameof(WarningCount));
        OnPropertyChanged(nameof(ErrorCount));
        OnPropertyChanged(nameof(HasErrors));
        OnPropertyChanged(nameof(HasWarnings));
    }

    private void CopyLogs()
    {
        var logs = FilteredLogs.ToList();
        var lines = logs.Select(l => $"[{l.Timestamp:HH:mm:ss.fff}] [{l.Level}] [{l.Source}] {l.Message}");
        var text = string.Join(Environment.NewLine, lines);

        try
        {
            System.Windows.Clipboard.SetText(text);
            _shellViewModel.AddLog(LogLevel.Information, $"Copied {logs.Count} log entries to clipboard");
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Failed to copy logs: {ex.Message}");
        }
    }

    private void ScrollToTop()
    {
        _userHasScrolled = false;
        OnPropertyChanged(nameof(UserHasScrolled));
        OnPropertyChanged(nameof(ShouldAutoScroll));
        // View binderà su questo evento per scrollare
    }

    private void ScrollToBottom()
    {
        _userHasScrolled = false;
        OnPropertyChanged(nameof(UserHasScrolled));
        OnPropertyChanged(nameof(ShouldAutoScroll));
        // View binderà su questo evento per scrollare
    }

    #endregion

    /// <summary>
    /// Rilascia le risorse.
    /// </summary>
    public void Dispose()
    {
        _shellViewModel.LogEntries.CollectionChanged -= OnLogEntriesChanged;
    }
}
