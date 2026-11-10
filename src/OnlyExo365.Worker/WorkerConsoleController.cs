using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using OnlyExo365.Contracts.Diagnostics;
using OnlyExo365.Contracts.Messages;

namespace OnlyExo365.Worker;

internal interface IWorkerConsoleNativeMethods
{
    IntPtr GetConsoleWindow();
    bool AllocConsole();
    bool FreeConsole();
    int GetLastError();
    bool ShowWindow(IntPtr windowHandle, int command);
    bool IsWindowVisible(IntPtr windowHandle);
    bool DisableCloseButton(IntPtr windowHandle);
}

internal sealed class WorkerConsoleNativeMethods : IWorkerConsoleNativeMethods
{
    public IntPtr GetConsoleWindow() => NativeMethods.GetConsoleWindow();
    public bool AllocConsole() => NativeMethods.AllocConsole();
    public bool FreeConsole() => NativeMethods.FreeConsole();
    public int GetLastError() => Marshal.GetLastWin32Error();
    public bool ShowWindow(IntPtr windowHandle, int command) => NativeMethods.ShowWindow(windowHandle, command);
    public bool IsWindowVisible(IntPtr windowHandle) => NativeMethods.IsWindowVisible(windowHandle);
    public bool DisableCloseButton(IntPtr windowHandle)
    {
        var systemMenuHandle = NativeMethods.GetSystemMenu(windowHandle, revert: false);
        if (systemMenuHandle == IntPtr.Zero)
        {
            return false;
        }

        var result = NativeMethods.EnableMenuItem(
            systemMenuHandle,
            NativeMethods.ScClose,
            NativeMethods.MfByCommand | NativeMethods.MfGrayed);
        NativeMethods.DrawMenuBar(windowHandle);
        return result != NativeMethods.EnableMenuItemFailed;
    }

    private static class NativeMethods
    {
        internal const uint ScClose = 0xF060;
        internal const uint MfByCommand = 0x00000000;
        internal const uint MfGrayed = 0x00000001;
        internal const uint EnableMenuItemFailed = 0xFFFFFFFF;

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern bool AllocConsole();

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern bool FreeConsole();

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr GetSystemMenu(IntPtr hWnd, bool revert);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern uint EnableMenuItem(IntPtr hMenu, uint uIDEnableItem, uint uEnable);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool DrawMenuBar(IntPtr hWnd);
    }
}

internal sealed class WorkerConsoleController
{
    private const int ErrorAccessDenied = 5;
    private const int SwHide = 0;
    private const int SwShow = 5;
    private const int VisibilityRetryCount = 10;
    private const int VisibilityRetryDelayMs = 20;

    private readonly IWorkerConsoleNativeMethods _nativeMethods;
    private readonly string _logDirectoryPath;
    private readonly Action _rebindConsoleStreams;
    private readonly Action<Action> _queueHistoryReplay;
    private readonly TextWriter? _historyOutput;
    private readonly object _sync = new();
    private bool _streamsRebound;
    private bool _historyReplayed;

    public WorkerConsoleController(
        IWorkerConsoleNativeMethods? nativeMethods = null,
        string? logDirectoryPath = null,
        Action? rebindConsoleStreams = null,
        Action<Action>? queueHistoryReplay = null,
        TextWriter? historyOutput = null)
    {
        _nativeMethods = nativeMethods ?? new WorkerConsoleNativeMethods();
        _logDirectoryPath = string.IsNullOrWhiteSpace(logDirectoryPath)
            ? PersistentLogWriter.GetDefaultLogDirectoryPath()
            : logDirectoryPath;
        _rebindConsoleStreams = rebindConsoleStreams ?? RebindConsoleStreams;
        _queueHistoryReplay = queueHistoryReplay ?? (action => Task.Run(action));
        _historyOutput = historyOutput;
    }

    public bool IsVisible
    {
        get
        {
            lock (_sync)
            {
                var windowHandle = _nativeMethods.GetConsoleWindow();
                return windowHandle != IntPtr.Zero && _nativeMethods.IsWindowVisible(windowHandle);
            }
        }
    }

    public SetWorkerConsoleVisibilityResponse SetVisibility(bool isVisible)
    {
        lock (_sync)
        {
            return isVisible ? ShowConsoleUnsafe() : HideConsoleUnsafe();
        }
    }

    private SetWorkerConsoleVisibilityResponse ShowConsoleUnsafe()
    {
        var windowHandle = _nativeMethods.GetConsoleWindow();
        if (windowHandle == IntPtr.Zero)
        {
            EnsureConsoleAllocated();
            RebindConsoleStreamsOnceUnsafe();
            windowHandle = WaitForConsoleWindowUnsafe();
            if (windowHandle == IntPtr.Zero)
            {
                throw new InvalidOperationException("Worker console handle not available after allocation.");
            }
        }

        if (!ShowWindowAndWaitForVisibilityUnsafe(windowHandle))
        {
            throw new InvalidOperationException("Unable to show the worker console window.");
        }

        DisableCloseButtonUnsafe(windowHandle);
        QueuePersistentWorkerLogReplayOnceUnsafe();

        return new SetWorkerConsoleVisibilityResponse
        {
            IsVisible = true,
            Message = "Worker console shown."
        };
    }

    private void EnsureConsoleAllocated()
    {
        if (_nativeMethods.AllocConsole())
        {
            return;
        }

        var errorCode = _nativeMethods.GetLastError();
        if (errorCode != ErrorAccessDenied)
        {
            throw new Win32Exception(errorCode, "Unable to allocate the worker console.");
        }

        EnsureNativeCallSucceeded(_nativeMethods.FreeConsole(), "Unable to detach the existing worker console.");

        if (_nativeMethods.AllocConsole())
        {
            return;
        }

        throw new Win32Exception(_nativeMethods.GetLastError(), "Unable to allocate the worker console.");
    }

    private SetWorkerConsoleVisibilityResponse HideConsoleUnsafe()
    {
        var windowHandle = _nativeMethods.GetConsoleWindow();
        if (windowHandle == IntPtr.Zero)
        {
            return new SetWorkerConsoleVisibilityResponse
            {
                IsVisible = false,
                Message = "Worker console already hidden."
            };
        }

        _nativeMethods.ShowWindow(windowHandle, SwHide);
        if (_nativeMethods.IsWindowVisible(windowHandle))
        {
            throw new InvalidOperationException("Unable to hide the worker console window.");
        }

        return new SetWorkerConsoleVisibilityResponse
        {
            IsVisible = false,
            Message = "Worker console hidden."
        };
    }

    private IntPtr WaitForConsoleWindowUnsafe()
    {
        for (var attempt = 0; attempt < VisibilityRetryCount; attempt++)
        {
            var windowHandle = _nativeMethods.GetConsoleWindow();
            if (windowHandle != IntPtr.Zero)
            {
                return windowHandle;
            }

            Thread.Sleep(VisibilityRetryDelayMs);
        }

        return _nativeMethods.GetConsoleWindow();
    }

    private bool ShowWindowAndWaitForVisibilityUnsafe(IntPtr windowHandle)
    {
        for (var attempt = 0; attempt < VisibilityRetryCount; attempt++)
        {
            _nativeMethods.ShowWindow(windowHandle, SwShow);
            if (_nativeMethods.IsWindowVisible(windowHandle))
            {
                return true;
            }

            Thread.Sleep(VisibilityRetryDelayMs);
        }

        return _nativeMethods.IsWindowVisible(windowHandle);
    }

    private void DisableCloseButtonUnsafe(IntPtr windowHandle)
    {
        try
        {
            _nativeMethods.DisableCloseButton(windowHandle);
        }
        catch (Win32Exception)
        {
        }
        catch (InvalidOperationException)
        {
        }
    }

    private void RebindConsoleStreamsOnceUnsafe()
    {
        if (_streamsRebound)
        {
            return;
        }

        _rebindConsoleStreams();
        _streamsRebound = true;
    }

    private void QueuePersistentWorkerLogReplayOnceUnsafe()
    {
        if (_historyReplayed)
        {
            return;
        }

        _historyReplayed = true;
        _queueHistoryReplay(() =>
        {
            try
            {
                foreach (var entry in ReadPersistentWorkerLogHistory())
                {
                    WriteHistoryEntry(entry);
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidOperationException or ObjectDisposedException)
            {
            }
        });
    }

    private IEnumerable<PersistentLogEntry> ReadPersistentWorkerLogHistory()
    {
        if (!Directory.Exists(_logDirectoryPath))
        {
            yield break;
        }

        foreach (var filePath in Directory.EnumerateFiles(_logDirectoryPath, "worker-*.log", SearchOption.TopDirectoryOnly)
                     .OrderBy(path => path, StringComparer.OrdinalIgnoreCase))
        {
            string[] lines;
            try
            {
                lines = File.ReadAllLines(filePath);
            }
            catch (IOException)
            {
                continue;
            }
            catch (UnauthorizedAccessException)
            {
                continue;
            }

            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                PersistentLogEntry? entry;
                try
                {
                    entry = JsonSerializer.Deserialize<PersistentLogEntry>(line);
                }
                catch (JsonException)
                {
                    continue;
                }

                if (entry != null)
                {
                    yield return entry;
                }
            }
        }
    }

    private void WriteHistoryEntry(PersistentLogEntry entry)
    {
        var output = _historyOutput ?? Console.Out;
        output.WriteLine(
            $"{entry.TimestampUtc.ToLocalTime():HH:mm:ss.fff} [{FormatLevel(entry.Level)}] [{entry.Source}] {entry.Message}");
    }

    private static string FormatLevel(LogLevel level) => level switch
    {
        LogLevel.Verbose => "VRB",
        LogLevel.Debug => "DBG",
        LogLevel.Information => "INF",
        LogLevel.Warning => "WRN",
        LogLevel.Error => "ERR",
        _ => "???"
    };

    private static void RebindConsoleStreams()
    {
#pragma warning disable CA2000
        var output = new StreamWriter(Console.OpenStandardOutput(), Encoding.UTF8) { AutoFlush = true };
        var error = new StreamWriter(Console.OpenStandardError(), Encoding.UTF8) { AutoFlush = true };
#pragma warning restore CA2000
        Console.SetOut(output);
        Console.SetError(error);
    }

    private static void EnsureNativeCallSucceeded(bool success, string message)
    {
        if (success)
        {
            return;
        }

        throw new Win32Exception(Marshal.GetLastWin32Error(), message);
    }
}

