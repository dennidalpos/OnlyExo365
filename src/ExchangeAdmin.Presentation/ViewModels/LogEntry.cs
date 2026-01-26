using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Presentation.ViewModels;

             
                            
              
public class LogEntry
{
    public DateTime Timestamp { get; init; } = DateTime.Now;
    public LogLevel Level { get; init; }
    public string Message { get; init; } = string.Empty;
    public string? Source { get; init; }

    public string FormattedTimestamp => Timestamp.ToString("HH:mm:ss.fff");

    public string LevelIcon => Level switch
    {
        LogLevel.Verbose => "V",
        LogLevel.Debug => "D",
        LogLevel.Information => "I",
        LogLevel.Warning => "W",
        LogLevel.Error => "E",
        _ => "?"
    };

    public string LevelColor => Level switch
    {
        LogLevel.Verbose => "#9D9D9D",
        LogLevel.Debug => "#569CD6",
        LogLevel.Information => "#4EC9B0",
        LogLevel.Warning => "#DCDCAA",
        LogLevel.Error => "#F14C4C",
        _ => "#CCCCCC"
    };
}
