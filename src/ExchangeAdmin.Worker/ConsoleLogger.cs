using System.Runtime.InteropServices;

namespace ExchangeAdmin.Worker;

public enum ConsoleLogLevel
{
    Verbose,
    Debug,
    Info,
    Warning,
    Error,
    Success
}

public static class ConsoleLogger
{
    private static readonly object _lock = new();
    private static bool _initialized;

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GetStdHandle(int nStdHandle);

    private const int STD_OUTPUT_HANDLE = -11;
    private const uint ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004;

    private static void EnsureInitialized()
    {
        if (_initialized) return;
        _initialized = true;

        try
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                var handle = GetStdHandle(STD_OUTPUT_HANDLE);
                if (handle != IntPtr.Zero && GetConsoleMode(handle, out uint mode))
                {
                    SetConsoleMode(handle, mode | ENABLE_VIRTUAL_TERMINAL_PROCESSING);
                }
            }
        }
        catch
        {
        }
    }

    public static void Log(ConsoleLogLevel level, string source, string message)
    {
        EnsureInitialized();
        lock (_lock)
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            var originalColor = Console.ForegroundColor;

            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.Write($"{timestamp} ");

            Console.ForegroundColor = GetLevelColor(level);
            Console.Write($"[{GetLevelPrefix(level)}] ");

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.Write($"[{source}] ");

            Console.ForegroundColor = GetMessageColor(level);
            Console.WriteLine(message);

            Console.ForegroundColor = originalColor;
        }
    }

    public static void Verbose(string source, string message) => Log(ConsoleLogLevel.Verbose, source, message);
    public static void Debug(string source, string message) => Log(ConsoleLogLevel.Debug, source, message);
    public static void Info(string source, string message) => Log(ConsoleLogLevel.Info, source, message);
    public static void Warning(string source, string message) => Log(ConsoleLogLevel.Warning, source, message);
    public static void Error(string source, string message) => Log(ConsoleLogLevel.Error, source, message);
    public static void Success(string source, string message) => Log(ConsoleLogLevel.Success, source, message);

    private static ConsoleColor GetLevelColor(ConsoleLogLevel level) => level switch
    {
        ConsoleLogLevel.Verbose => ConsoleColor.DarkGray,
        ConsoleLogLevel.Debug => ConsoleColor.Blue,
        ConsoleLogLevel.Info => ConsoleColor.White,
        ConsoleLogLevel.Warning => ConsoleColor.Yellow,
        ConsoleLogLevel.Error => ConsoleColor.Red,
        ConsoleLogLevel.Success => ConsoleColor.Green,
        _ => ConsoleColor.Gray
    };

    private static ConsoleColor GetMessageColor(ConsoleLogLevel level) => level switch
    {
        ConsoleLogLevel.Verbose => ConsoleColor.DarkGray,
        ConsoleLogLevel.Debug => ConsoleColor.Gray,
        ConsoleLogLevel.Info => ConsoleColor.White,
        ConsoleLogLevel.Warning => ConsoleColor.Yellow,
        ConsoleLogLevel.Error => ConsoleColor.Red,
        ConsoleLogLevel.Success => ConsoleColor.Green,
        _ => ConsoleColor.Gray
    };

    private static string GetLevelPrefix(ConsoleLogLevel level) => level switch
    {
        ConsoleLogLevel.Verbose => "VRB",
        ConsoleLogLevel.Debug => "DBG",
        ConsoleLogLevel.Info => "INF",
        ConsoleLogLevel.Warning => "WRN",
        ConsoleLogLevel.Error => "ERR",
        ConsoleLogLevel.Success => "OK ",
        _ => "???"
    };
}
