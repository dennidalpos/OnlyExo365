using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
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
}

internal sealed class WorkerConsoleNativeMethods : IWorkerConsoleNativeMethods
{
    public IntPtr GetConsoleWindow() => NativeMethods.GetConsoleWindow();
    public bool AllocConsole() => NativeMethods.AllocConsole();
    public bool FreeConsole() => NativeMethods.FreeConsole();
    public int GetLastError() => Marshal.GetLastWin32Error();
    public bool ShowWindow(IntPtr windowHandle, int command) => NativeMethods.ShowWindow(windowHandle, command);
    public bool IsWindowVisible(IntPtr windowHandle) => NativeMethods.IsWindowVisible(windowHandle);

    private static class NativeMethods
    {
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
    }
}

internal sealed class WorkerConsoleController
{
    private const int ErrorAccessDenied = 5;
    private const int SwHide = 0;
    private const int SwShow = 5;

    private readonly IWorkerConsoleNativeMethods _nativeMethods;
    private readonly object _sync = new();
    private bool _allocatedByController;

    public WorkerConsoleController(IWorkerConsoleNativeMethods? nativeMethods = null)
    {
        _nativeMethods = nativeMethods ?? new WorkerConsoleNativeMethods();
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
            _allocatedByController = true;
            RebindConsoleStreams();
            windowHandle = _nativeMethods.GetConsoleWindow();
            if (windowHandle == IntPtr.Zero)
            {
                throw new InvalidOperationException("Worker console handle not available after allocation.");
            }
        }

        _nativeMethods.ShowWindow(windowHandle, SwShow);
        if (!_nativeMethods.IsWindowVisible(windowHandle))
        {
            throw new InvalidOperationException("Unable to show the worker console window.");
        }

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
            _allocatedByController = false;
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

        if (_allocatedByController)
        {
            EnsureNativeCallSucceeded(_nativeMethods.FreeConsole(), "Unable to release the worker console.");
            _allocatedByController = false;
        }

        return new SetWorkerConsoleVisibilityResponse
        {
            IsVisible = false,
            Message = "Worker console hidden."
        };
    }

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

