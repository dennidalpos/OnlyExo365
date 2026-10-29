using OnlyExo365.Shell.Ipc;

namespace OnlyExo365.Tests;

public class WorkerSupervisorTests
{
    [Fact]
    public void ResolveWorkerPath_UsesBaseDirectoryForRelativeConfiguredPath()
    {
        var baseDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));

        var resolvedPath = WorkerSupervisor.ResolveWorkerPath(
            configuredWorkerPath: Path.Combine("workers", "OnlyExo365.Worker.exe"),
            baseDirectory: baseDirectory);

        Assert.Equal(
            Path.GetFullPath(Path.Combine(baseDirectory, "workers", "OnlyExo365.Worker.exe")),
            resolvedPath);
    }

    [Fact]
    public void ResolveWorkerPath_UsesDefaultWorkerExecutableNameWhenConfigurationIsMissing()
    {
        var baseDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));

        var resolvedPath = WorkerSupervisor.ResolveWorkerPath(
            configuredWorkerPath: null,
            baseDirectory: baseDirectory);

        Assert.Equal(
            Path.GetFullPath(Path.Combine(baseDirectory, "OnlyExo365.Worker.exe")),
            resolvedPath);
    }

    [Fact]
    public async Task RestartWithinBudgetAsync_UsesRestartActionWhileBudgetAvailable()
    {
        var invocations = 0;

        var result = await WorkerSupervisor.RestartWithinBudgetAsync(
            restartCount: 1,
            maxRestartAttempts: 3,
            restartAction: _ =>
            {
                invocations++;
                return Task.FromResult(true);
            });

        Assert.True(result);
        Assert.Equal(1, invocations);
    }

    [Fact]
    public async Task RestartWithinBudgetAsync_SkipsRestartActionWhenBudgetIsExhausted()
    {
        var invocations = 0;

        var result = await WorkerSupervisor.RestartWithinBudgetAsync(
            restartCount: 3,
            maxRestartAttempts: 3,
            restartAction: _ =>
            {
                invocations++;
                return Task.FromResult(true);
            });

        Assert.False(result);
        Assert.Equal(0, invocations);
    }

    [Fact]
    public async Task GetStatus_ReflectsTrackedWorkerConsoleVisibility()
    {
        await using var supervisor = new WorkerSupervisor();

        supervisor.SetConsoleVisibility(true);

        var status = supervisor.GetStatus();

        Assert.True(status.IsConsoleVisible);
    }
}

