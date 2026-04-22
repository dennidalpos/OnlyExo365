using System.Windows;
using OnlyExo365.Contracts;
using OnlyExo365.Shell.Views;

namespace OnlyExo365.Shell.Bootstrap;

internal static class AppCompositionRoot
{
    public static AppRuntimeContext Create(ExchangeOnlineConfiguration exchangeConfiguration)
    {
        var modules = AppModuleFactory.Create(exchangeConfiguration);
        AppShellModuleRegistrar.Configure(modules);

        var pageLoadCoordinator = new AppPageLoadCoordinator(
            modules.ShellViewModel,
            modules.NavigationService,
            AppPageLoaderCatalog.Create(modules));

        var mainWindow = new MainWindow
        {
            DataContext = modules.ShellViewModel
        };

        var runtimeContext = new AppRuntimeContext
        {
            WorkerService = modules.WorkerService,
            ShellViewModel = modules.ShellViewModel,
            NavigationService = modules.NavigationService,
            MainWindow = mainWindow,
            PageLoadCoordinator = pageLoadCoordinator,
            CatalogUpdateService = modules.LicenseCatalogUpdateService
        };

        mainWindow.ShutdownHandler = runtimeContext.RequestShutdownAsync;
        return runtimeContext;
    }
}

