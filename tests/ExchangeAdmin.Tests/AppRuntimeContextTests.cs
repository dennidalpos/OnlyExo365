using System.IO;
using ExchangeAdmin.Presentation.Bootstrap;
using ExchangeAdmin.Presentation.Services;
using ExchangeAdmin.Presentation.ViewModels;

namespace ExchangeAdmin.Tests;

public sealed class AppRuntimeContextTests
{
    private static LicenseCatalogUpdateService CreateNoOpCatalogService()
    {
        var cfg = new LicenseCatalogConfiguration
        {
            AutoUpdateMode = CatalogAutoUpdateMode.Disabled,
            CheckOnStartup = false,
            LocalCachePath = Path.Combine(Path.GetTempPath(), "ExchangeAdmin.Tests.Runtime", Guid.NewGuid().ToString("N"))
        };
#pragma warning disable CA2000 // LicenseCatalogUpdateService owns and disposes the downloader.
        var downloader = new LicenseCatalogDownloader(cfg);
        try
        {
            return new LicenseCatalogUpdateService(
                cfg,
                new LicenseCatalogFileStore(cfg),
                downloader,
                new PresentationSkuNameResolver());
        }
        catch
        {
            downloader.Dispose();
            throw;
        }
#pragma warning restore CA2000
    }

    [Fact]
    public async Task RequestShutdownAsync_IsIdempotentAcrossMultipleCalls()
    {
        var workerService = new TrackingWorkerService();
        var navigationService = new NavigationService();
        var shellViewModel = new ShellViewModel(workerService, navigationService);
        using var catalogService = CreateNoOpCatalogService();
        var runtimeContext = new AppRuntimeContext
        {
            WorkerService = workerService,
            ShellViewModel = shellViewModel,
            NavigationService = navigationService,
            MainWindow = null!,
            PageLoadCoordinator = new AppPageLoadCoordinator(
                shellViewModel,
                navigationService,
                new Dictionary<NavigationPage, Func<Task>>()),
            CatalogUpdateService = catalogService
        };

        await Task.WhenAll(runtimeContext.RequestShutdownAsync(), runtimeContext.RequestShutdownAsync());

        Assert.Equal(1, workerService.StopCalls);
        Assert.Equal(1, workerService.DisposeCalls);

        await runtimeContext.DisposeAsync();

        Assert.Equal(1, workerService.StopCalls);
        Assert.Equal(1, workerService.DisposeCalls);
    }

    private sealed class TrackingWorkerService : TestWorkerServiceBase, IAsyncDisposable
    {
        public int StopCalls { get; private set; }
        public int DisposeCalls { get; private set; }

        public override Task StopWorkerAsync()
        {
            StopCalls++;
            return Task.CompletedTask;
        }

        public ValueTask DisposeAsync()
        {
            DisposeCalls++;
            return ValueTask.CompletedTask;
        }
    }
}
