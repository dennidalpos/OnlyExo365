using System.ComponentModel;
using OnlyExo365.Worker;

namespace OnlyExo365.Tests;

public sealed class WorkerConsoleControllerTests
{
    [Fact]
    public void SetVisibility_ShowAllocatesAndShowsConsoleWhenMissing()
    {
        var native = new FakeWorkerConsoleNativeMethods();
        var controller = new WorkerConsoleController(native);

        var response = controller.SetVisibility(true);

        Assert.True(response.IsVisible);
        Assert.True(native.AllocConsoleCalled);
        Assert.Equal(1, native.ShowWindowCalls);
        Assert.True(controller.IsVisible);
    }

    [Fact]
    public void SetVisibility_HideReleasesConsoleAllocatedByController()
    {
        var native = new FakeWorkerConsoleNativeMethods();
        var controller = new WorkerConsoleController(native);
        controller.SetVisibility(true);

        var response = controller.SetVisibility(false);

        Assert.False(response.IsVisible);
        Assert.True(native.FreeConsoleCalled);
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
        var controller = new WorkerConsoleController(native);

        var response = controller.SetVisibility(true);

        Assert.True(response.IsVisible);
        Assert.False(native.AllocConsoleCalled);
        Assert.Equal(1, native.ShowWindowCalls);
    }

    [Fact]
    public void SetVisibility_HideIsIdempotentWhenConsoleAlreadyHidden()
    {
        var native = new FakeWorkerConsoleNativeMethods();
        var controller = new WorkerConsoleController(native);

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
        var controller = new WorkerConsoleController(native);

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
        var controller = new WorkerConsoleController(native);

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
        var controller = new WorkerConsoleController(native);

        var response = controller.SetVisibility(true);

        Assert.True(response.IsVisible);
        Assert.True(controller.IsVisible);
    }

    private sealed class FakeWorkerConsoleNativeMethods : IWorkerConsoleNativeMethods
    {
        public bool HasConsoleWindow { get; set; }
        public bool IsVisible { get; set; }
        public bool AllocConsoleResult { get; set; } = true;
        public bool FreeConsoleResult { get; set; } = true;
        public bool ShowWindowResult { get; set; } = true;
        public bool AllocConsoleCalled { get; private set; }
        public int AllocConsoleCalls { get; private set; }
        public bool FreeConsoleCalled { get; private set; }
        public int ShowWindowCalls { get; private set; }
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
    }
}

