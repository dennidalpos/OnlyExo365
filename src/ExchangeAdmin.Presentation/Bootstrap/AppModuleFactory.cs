using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts;
using ExchangeAdmin.Infrastructure.Ipc;
using ExchangeAdmin.Presentation.Localization;
using ExchangeAdmin.Presentation.Services;
using ExchangeAdmin.Presentation.ViewModels;

namespace ExchangeAdmin.Presentation.Bootstrap;

internal static class AppModuleFactory
{
    public static AppModuleCatalog Create(ExchangeOnlineConfiguration exchangeConfiguration)
    {
        // Load and apply the persisted locale preference before any VM is constructed
        // so that all VMs see the correct locale from the start.
        var preferencesService = new UserPreferencesService();
        var savedLocale = preferencesService.LoadLocale();
        if (!string.IsNullOrEmpty(savedLocale))
        {
            LocalizationService.Instance.SetLocale(savedLocale);
        }

        var languageSelectionViewModel = new LanguageSelectionViewModel(preferencesService);

        var workerOptions = new WorkerSupervisorOptions
        {
            WorkerPath = "ExchangeAdmin.Worker.exe",
            MaxRestartAttempts = 3,
            ExchangeConfiguration = exchangeConfiguration
        };

        var workerService = new WorkerService(workerOptions);
        var navigationService = new NavigationService();
        var cacheService = new CacheService();
        var interactiveExchangeBootstrapService = new InteractiveExchangeBootstrapService(exchangeConfiguration);
        var shellViewModel = new ShellViewModel(
            workerService,
            navigationService,
            exchangeConfiguration,
            interactiveExchangeBootstrapService);

        // License catalog services (run in Presentation process; no IPC needed).
        var catalogConfig = ExchangeConfigurationLoader.LoadLicenseCatalogConfiguration();
        var catalogFileStore = new LicenseCatalogFileStore(catalogConfig);
        var catalogDownloader = new LicenseCatalogDownloader(catalogConfig);
        var skuNameResolver = new PresentationSkuNameResolver();
        var catalogUpdateService = new LicenseCatalogUpdateService(
            catalogConfig, catalogFileStore, catalogDownloader, skuNameResolver);
        var catalogViewModel = new LicenseCatalogViewModel(catalogUpdateService, shellViewModel);

        return new AppModuleCatalog
        {
            WorkerService = workerService,
            NavigationService = navigationService,
            LicenseCatalogUpdateService = catalogUpdateService,
            ShellViewModel = shellViewModel,
            LanguageSelectionViewModel = languageSelectionViewModel,
            DashboardViewModel = new DashboardViewModel(workerService, navigationService, shellViewModel, cacheService),
            ContactsViewModel = new ContactsViewModel(workerService, shellViewModel),
            ResourcesViewModel = new ResourcesViewModel(workerService, shellViewModel),
            PublicFoldersViewModel = new PublicFoldersViewModel(workerService, shellViewModel),
            MobileDevicesViewModel = new MobileDevicesViewModel(workerService, shellViewModel),
            MigrationViewModel = new MigrationViewModel(workerService, shellViewModel),
            PermissionsViewModel = new PermissionsViewModel(workerService, shellViewModel),
            MailboxListViewModel = new MailboxListViewModel(workerService, navigationService, shellViewModel),
            DeletedMailboxesViewModel = new DeletedMailboxesViewModel(workerService, shellViewModel),
            MailboxDetailsViewModel = new MailboxDetailsViewModel(workerService, navigationService, shellViewModel, cacheService),
            MailboxSpaceViewModel = new MailboxSpaceViewModel(workerService, navigationService, shellViewModel),
            MailboxAccessReportViewModel = new MailboxAccessReportViewModel(workerService, shellViewModel),
            DistributionListViewModel = new DistributionListViewModel(workerService, navigationService, shellViewModel),
            LogsViewModel = new LogsViewModel(shellViewModel),
            ToolsViewModel = new ToolsViewModel(workerService, shellViewModel, exchangeConfiguration, catalogViewModel),
            MessageTraceViewModel = new MessageTraceViewModel(workerService, shellViewModel),
            ComplianceViewModel = new ComplianceViewModel(workerService, shellViewModel),
            MailSecurityViewModel = new MailSecurityViewModel(workerService, shellViewModel),
            MailFlowViewModel = new MailFlowViewModel(workerService, shellViewModel)
        };
    }
}
