using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Windows.Input;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;

namespace ExchangeAdmin.Presentation.ViewModels;

             
                                                                                 
              
public sealed class LogsViewModel : ViewModelBase, IDisposable
{
    private readonly ShellViewModel _shellViewModel;
    private readonly object _filterLock = new();
    private readonly DebounceHelper _refreshDebounce = new();

    private LogLevel _filterLevel = LogLevel.Verbose;
    private string? _searchFilter;
    private bool _autoScroll = true;
    private bool _userHasScrolled;
    private bool _isRefreshing;

                                                     
    private List<LogEntry>? _cachedFilteredLogs;
    private bool _filterCacheInvalid = true;

                                            
    private DateTime _lastRefreshTime = DateTime.MinValue;
    private static readonly TimeSpan MinRefreshInterval = TimeSpan.FromMilliseconds(100);

                 
                                          
                  
                                                                                
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

                                                    
        _shellViewModel.LogEntries.CollectionChanged += OnLogEntriesChanged;
    }

    #region Properties

                 
                                                        
                  
    public ObservableCollection<LogEntry> LogEntries => _shellViewModel.LogEntries;

                 
                                              
                  
    public LogLevel FilterLevel
    {
        get => _filterLevel;
        set
        {
            if (SetProperty(ref _filterLevel, value))
            {
                InvalidateFilterCache();
                RequestFilterRefresh(immediate: true);
            }
        }
    }

                 
                                   
                  
    public string? SearchFilter
    {
        get => _searchFilter;
        set
        {
            if (SetProperty(ref _searchFilter, value))
            {
                InvalidateFilterCache();
                RequestFilterRefresh(immediate: false);
            }
        }
    }

                 
                                                                                           
                  
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

                 
                                                                                             
                  
    public bool UserHasScrolled
    {
        get => _userHasScrolled;
        set => SetProperty(ref _userHasScrolled, value);
    }

                 
                                                                                                   
                  
    public bool ShouldAutoScroll => AutoScroll && !_userHasScrolled;

                 
                                 
                  
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

                 
                                   
                  
    public int TotalCount => LogEntries.Count;

                 
                                        
                  
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

                                                         
    public int VerboseCount => CountByLevel(LogLevel.Verbose);

                                                       
    public int DebugCount => CountByLevel(LogLevel.Debug);

                                                             
    public int InfoCount => CountByLevel(LogLevel.Information);

                                                         
    public int WarningCount => CountByLevel(LogLevel.Warning);

                                                       
    public int ErrorCount => CountByLevel(LogLevel.Error);

                 
                                        
                  
    public bool HasErrors => ErrorCount > 0;

                 
                                  
                  
    public bool HasWarnings => WarningCount > 0;

    #endregion

    #region Commands

                                                
    public ICommand ClearLogsCommand { get; }

                                                              
    public ICommand CopyLogsCommand { get; }

                                                      
    public ICommand ScrollToTopCommand { get; }

                                                     
    public ICommand ScrollToBottomCommand { get; }

    #endregion

    #region Methods

                 
                                          
                  
    public void Refresh()
    {
        InvalidateFilterCache();
        NotifyFilteredLogsChanged();
        NotifyCountsChanged();
    }

                 
                                                       
                  
    public void NotifyUserScrolled()
    {
        if (AutoScroll)
        {
            _userHasScrolled = true;
            OnPropertyChanged(nameof(UserHasScrolled));
            OnPropertyChanged(nameof(ShouldAutoScroll));
        }
    }

                 
                                                                      
                  
    public void ResetUserScroll()
    {
        _userHasScrolled = false;
        OnPropertyChanged(nameof(UserHasScrolled));
        OnPropertyChanged(nameof(ShouldAutoScroll));
    }

    private void OnLogEntriesChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        InvalidateFilterCache();

                                                                              
        var now = DateTime.UtcNow;
        if (now - _lastRefreshTime >= MinRefreshInterval)
        {
            _lastRefreshTime = now;

            RunOnUiThread(() =>
            {
                RequestFilterRefresh(immediate: true);
                NotifyCountsChanged();
            });
        }
        else
        {
            var delayMs = (int)Math.Max(1, (MinRefreshInterval - (now - _lastRefreshTime)).TotalMilliseconds);
            _refreshDebounce.Debounce(() =>
            {
                _lastRefreshTime = DateTime.UtcNow;
                RunOnUiThread(() =>
                {
                    RequestFilterRefresh(immediate: true);
                    NotifyCountsChanged();
                });
            }, delayMs);
        }
    }


    private void RequestFilterRefresh(bool immediate)
    {
        if (immediate)
        {
            NotifyFilteredLogsChanged();
            return;
        }

        _refreshDebounce.Debounce(() => RunOnUiThread(NotifyFilteredLogsChanged), 200);
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
                           
        _cachedFilteredLogs = new List<LogEntry>();
        var searchLower = string.IsNullOrWhiteSpace(_searchFilter) ? null : _searchFilter.Trim().ToLowerInvariant();

        foreach (var log in LogEntries)
        {
                                 
            if (log.Level < _filterLevel)
            {
                continue;
            }

                                 
            if (!string.IsNullOrEmpty(searchLower))
            {
                var messageContains = (log.Message ?? string.Empty).ToLowerInvariant().Contains(searchLower);
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
                                                       
    }

    private void ScrollToBottom()
    {
        _userHasScrolled = false;
        OnPropertyChanged(nameof(UserHasScrolled));
        OnPropertyChanged(nameof(ShouldAutoScroll));
                                                       
    }

    #endregion

                 
                            
                  
    public void Dispose()
    {
        _shellViewModel.LogEntries.CollectionChanged -= OnLogEntriesChanged;
        _refreshDebounce.Dispose();
    }
}
