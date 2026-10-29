using System.ComponentModel;
using System.Text.Json;
using OnlyExo365.Contracts.Diagnostics;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Worker;

namespace OnlyExo365.Tests;

public sealed class WorkerConsoleControllerTests : IDisposable
{
    private readonly string _tempDirectory = Path.Combine(Path.GetTempPath(), "OnlyExo365.Tests", Guid.NewGuid().ToString("N"));
    private int _rebindCalls;

    public WorkerConsoleControllerTests()
    {
        Directory.CreateDirectory(_tempDirectory);
    }

    [Fact]
    public void SetVisibility_ShowAllocatesAndShowsConsoleWhenMissing()
    {
        var native = new FakeWorkerConsoleNativeMethods();
        var controller = CreateController(native);

        var response = controller.SetVisibility(true);

        Assert.True(response.IsVisible);
        Assert.True(native.AllocConsoleCalled);
        Assert.Equal(1, native.ShowWindowCalls);
        Assert.Equal(1, native.DisableCloseButtonCalls);
        Assert.Equal(1, _rebindCalls);
        Assert.True(controller.IsVisible);
    }

    [Fact]
    public void SetVisibility_HidePreservesConsoleAllocatedByController()
    {
        var native = new FakeWorkerConsoleNativeMethods();
        var controller = CreateController(native);
        controller.SetVisibility(true);

        var response = controller.SetVisibility(false);

        Assert.False(response.IsVisible);
        Assert.False(native.FreeConsoleCalled);
        Assert.True(native.HasConsoleWindow);
        Assert.False(controller.IsVisible);
    }

    [Fact]
    public void SetVisibility_ShowIsIdempotentWhenConsoleAlreadyVisible()
    {
        var native = new FakeWorkerConsoleNativeMethods
        {
            HasConsoleWindow = true,
            IsVisible = true
        };
        var controller = CreateController(native);

        var response = controller.SetVisibility(true);

        Assert.True(response.IsVisible);
        Assert.False(native.AllocConsoleCalled);
        Assert.Equal(1, native.ShowWindowCalls);
        Assert.Equal(1, native.DisableCloseButtonCalls);
    }

    [Fact]
    public void SetVisibility_HideIsIdempotentWhenConsoleAlreadyHidden()
    {
        var native = new FakeWorkerConsoleNativeMethods();
        var controller = CreateController(native);

        var response = controller.SetVisibility(false);

        Assert.False(response.IsVisible);
        Assert.False(native.FreeConsoleCalled);
        Assert.False(controller.IsVisible);
    }

    [Fact]
    public void SetVisibility_ThrowsControlledErrorWhenNativeAllocationFails()
    {
        var native = new FakeWorkerConsoleNativeMethods
        {
            AllocConsoleResult = false,
            LastError = 6
        };
        var controller = CreateController(native);

        var exception = Assert.Throws<Win32Exception>(() => controller.SetVisibility(true));

        Assert.Contains("allocate", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetVisibility_ShowDetachesAndReallocatesWhenAlreadyAttachedWithoutConsoleWindow()
    {
        var native = new FakeWorkerConsoleNativeMethods
        {
            AllocConsoleResult = false,
            LastError = 5,
            FailFirstAllocOnly = true
        };
        var controller = CreateController(native);

        var response = controller.SetVisibility(true);

        Assert.True(response.IsVisible);
        Assert.Equal(2, native.AllocConsoleCalls);
        Assert.True(native.FreeConsoleCalled);
        Assert.True(controller.IsVisible);
    }

    [Fact]
    public void SetVisibility_DoesNotTreatShowWindowReturnValueAsFailure()
    {
        var native = new FakeWorkerConsoleNativeMethods
        {
            ShowWindowResult = false
        };
        var controller = CreateController(native);

        var response = controller.SetVisibility(true);

        Assert.True(response.IsVisible);
        Assert.True(controller.IsVisible);
    }

    [Fact]
    public void SetVisibility_DoesNotFailWhenCloseButtonCannotBeDisabled()
    {
        var native = new FakeWorkerConsoleNativeMethods
        {
            DisableCloseButtonResult = false
        };
        var controller = CreateController(native);

        var response = controller.SetVisibility(true);

        Assert.True(response.IsVisible);
        Assert.Equal(1, native.DisableCloseButtonCalls);
        Assert.True(controller.IsVisible);
    }

    [Fact]
    public void SetVisibility_RepeatedShowHideShowDoesNotReallocateOrReleaseConsole()
    {
        var native = new FakeWorkerConsoleNativeMethods();
        var controller = CreateController(native);

        controller.SetVisibility(true);
        controller.SetVisibility(false);
        var response = controller.SetVisibility(true);

        Assert.True(response.IsVisible);
        Assert.Equal(1, native.AllocConsoleCalls);
        Assert.False(native.FreeConsoleCalled);
        Assert.Equal(3, native.ShowWindowCalls);
        Assert.Equal(2, native.DisableCloseButtonCalls);
        Assert.Equal(1, _rebindCalls);
        Assert.True(controller.IsVisible);
    }

    [Fact]
    public void SetVisibility_ReplaysPersistentWorkerHistoryOnce()
    {
        WriteWorkerLog(
            "worker-20260427.log",
            new PersistentLogEntry
            {
                TimestampUtc = new DateTime(2026, 4, 27, 8, 15, 0, DateTimeKind.Utc),
                Level = LogLevel.Information,
                Component = "worker",
                Source = "IPC",
                Message = "Historic startup",
                ProcessId = 100
            });
        var native = new FakeWorkerConsoleNativeMethods();
        using var output = new StringWriter();
        var controller = CreateController(native, output);

        controller.SetVisibility(true);
        var firstReplay = output.ToString();
        controller.SetVisibility(false);
        controller.SetVisibility(true);

        Assert.Contains("Historic startup", firstReplay, StringComparison.Ordinal);
        Assert.Equal(firstReplay, output.ToString());
    }

    [Fact]
    public void SetVisibility_SkipsMalformedPersistentWorkerLogLines()
    {
        File.WriteAllLines(
            Path.Combine(_tempDirectory, "worker-20260428.log"),
            [
                "not-json",
                JsonSerializer.Serialize(new PersistentLogEntry
                {
                    TimestampUtc = new DateTime(2026, 4, 28, 9, 0, 0, DateTimeKind.Utc),
                    Level = LogLevel.Warning,
                    Component = "worker",
                    Source = "IPC",
                    Message = "Recovered after malformed line",
                    ProcessId = 101
                })
            ]);
        var native = new FakeWorkerConsoleNativeMethods();
        using var output = new StringWriter();
        var controller = CreateController(native, output);

        var response = controller.SetVisibility(true);

        Assert.True(response.IsVisible);
        Assert.DoesNotContain("not-json", output.ToString(), StringComparison.Ordinal);
        Assert.Contains("Recovered after malformed line", output.ToString(), StringComparison.Ordinal);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDirectory))
            {
                Directory.Delete(_tempDirectory, recursive: true);
            }
        }
        catch
        {
        }
    }

    private WorkerConsoleController CreateController(FakeWorkerConsoleNativeMethods native, TextWriter? historyOutput = null)
        => new(
            native,
            _tempDirectory,
            rebindConsoleStreams: () => _rebindCalls++,
            queueHistoryReplay: action => action(),
            historyOutput: historyOutput);

    private void WriteWorkerLog(string fileName, params PersistentLogEntry[] entries)
    {
        var lines = entries.Select(entry => JsonSerializer.Serialize(entry));
        File.WriteAllLines(Path.Combine(_tempDirectory, fileName), lines);
    }

    private sealed class FakeWorkerConsoleNativeMethods : IWorkerConsoleNativeMethods
    {
        public bool HasConsoleWindow { get; set; }
        public bool IsVisible { get; set; }
        public bool AllocConsoleResult { get; set; } = true;
        public bool FreeConsoleResult { get; set; } = true;
        public bool ShowWindowResult { get; set; } = true;
        public bool DisableCloseButtonResult { get; set; } = true;
        public bool AllocConsoleCalled { get; private set; }
        public int AllocConsoleCalls { get; private set; }
        public bool FreeConsoleCalled { get; private set; }
        public int ShowWindowCalls { get; private set; }
        public int DisableCloseButtonCalls { get; private set; }
        public int LastError { get; set; }
        public bool FailFirstAllocOnly { get; set; }

        public IntPtr GetConsoleWindow() => HasConsoleWindow ? new IntPtr(42) : IntPtr.Zero;

        public bool AllocConsole()
        {
            AllocConsoleCalled = true;
            AllocConsoleCalls++;
            var shouldSucceed = AllocConsoleResult;
            if (FailFirstAllocOnly && AllocConsoleCalls > 1)
            {
                shouldSucceed = true;
                LastError = 0;
            }

            if (shouldSucceed)
            {
                HasConsoleWindow = true;
                IsVisible = true;
            }

            return shouldSucceed;
        }

        public bool FreeConsole()
        {
            FreeConsoleCalled = true;
            if (FreeConsoleResult)
            {
                HasConsoleWindow = false;
                IsVisible = false;
            }

            return FreeConsoleResult;
        }

        public int GetLastError() => LastError;

        public bool ShowWindow(IntPtr windowHandle, int command)
        {
            ShowWindowCalls++;
            if (ShowWindowResult)
            {
                IsVisible = command != 0;
            }

            return ShowWindowResult;
        }

        public bool IsWindowVisible(IntPtr windowHandle) => IsVisible;

        public bool DisableCloseButton(IntPtr windowHandle)
        {
            DisableCloseButtonCalls++;
            return DisableCloseButtonResult;
        }
    }
}

