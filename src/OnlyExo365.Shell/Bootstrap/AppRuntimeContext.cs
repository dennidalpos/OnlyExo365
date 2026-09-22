using System.Windows;
using OnlyExo365.Shell.Services;
using OnlyExo365.Shell.ViewModels;

namespace OnlyExo365.Shell.Bootstrap;

internal sealed class AppRuntimeContext : IAsyncDisposable
{
    private bool _disposed;
    private readonly object _shutdownLock = new();
    private Task? _shutdownTask;

    public required IWorkerService WorkerService { get; init; }
    public required ShellViewModel ShellViewModel { get; init; }
    public required NavigationService NavigationService { get; init; }
    public required Window MainWindow { get; init; }
    public required AppPageLoadCoordinator PageLoadCoordinator { get; init; }
    public required LicenseCatalogUpdateService CatalogUpdateService { get; init; }

    public async Task StartAsync()
    {
        await ShellViewModel.StartWorkerOnStartupAsync();

        // Non-blocking background catalog initialization.
        _ = Task.Run(() => CatalogUpdateService.InitializeAsync());
    }

    public Task RequestShutdownAsync()
    {
        lock (_shutdownLock)
        {
            _shutdownTask ??= ShutdownCoreAsync();
            return _shutdownTask;
        }
    }

    public async ValueTask DisposeAsync()
    {
        await RequestShutdownAsync();
    }

    private async Task ShutdownCoreAsync()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        PageLoadCoordinator.Dispose();
        CatalogUpdateService.Dispose();
        ShellViewModel.Dispose();
        await WorkerService.StopWorkerAsync();

        if (WorkerService is IAsyncDisposable asyncDisposable)
        {
            await asyncDisposable.DisposeAsync();
        }
    }
}

