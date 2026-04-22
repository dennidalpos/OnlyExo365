using System.IO;
using System.Reflection;
using OnlyExo365.Contracts;
using OnlyExo365.Shell.Bootstrap;
using OnlyExo365.Shell.Services;
using OnlyExo365.Shell.ViewModels;

namespace OnlyExo365.Tests;

public sealed class AppBootstrapCompositionTests
{
    [Fact]
    public void AppShellModuleRegistrar_AssignsAllShellModules()
    {
        var modules = CreateModules();

        AppShellModuleRegistrar.AssignModules(modules);

        Assert.Same(modules.DashboardViewModel, modules.ShellViewModel.Dashboard);
        Assert.Same(modules.ContactsViewModel, modules.ShellViewModel.Contacts);
        Assert.Same(modules.ResourcesViewModel, modules.ShellViewModel.Resources);
        Assert.Same(modules.PublicFoldersViewModel, modules.ShellViewModel.PublicFolders);
        Assert.Same(modules.MobileDevicesViewModel, modules.ShellViewModel.MobileDevices);
        Assert.Same(modules.MigrationViewModel, modules.ShellViewModel.Migration);
        Assert.Same(modules.PermissionsViewModel, modules.ShellViewModel.Permissions);
        Assert.Same(modules.MailboxListViewModel, modules.ShellViewModel.Mailboxes);
        Assert.Same(modules.DeletedMailboxesViewModel, modules.ShellViewModel.DeletedMailboxes);
        Assert.Same(modules.MailboxDetailsViewModel, modules.ShellViewModel.MailboxDetails);
        Assert.Same(modules.MailboxSpaceViewModel, modules.ShellViewModel.MailboxSpace);
        Assert.Same(modules.MailboxAccessReportViewModel, modules.ShellViewModel.MailboxAccessReport);
        Assert.Same(modules.DistributionListViewModel, modules.ShellViewModel.DistributionLists);
        Assert.Same(modules.LogsViewModel, modules.ShellViewModel.Logs);
        Assert.Same(modules.ToolsViewModel, modules.ShellViewModel.Tools);
        Assert.Same(modules.MessageTraceViewModel, modules.ShellViewModel.MessageTrace);
        Assert.Same(modules.ComplianceViewModel, modules.ShellViewModel.Compliance);
        Assert.Same(modules.MailSecurityViewModel, modules.ShellViewModel.MailSecurity);
        Assert.Same(modules.MailFlowViewModel, modules.ShellViewModel.MailFlow);
        Assert.Same(modules.LanguageSelectionViewModel, modules.ShellViewModel.LanguageSelection);
    }

    [Fact]
    public void AppPageLoaderCatalog_ReturnsExpectedNavigationPages()
    {
        var modules = CreateModules();

        var loaders = AppPageLoaderCatalog.Create(modules);
        var expectedPages = Enum.GetValues<NavigationPage>()
            .Except(new[]
            {
                NavigationPage.MailboxSpace,
                NavigationPage.MailboxAccessReport
            })
            .OrderBy(page => page)
            .ToArray();

        Assert.Equal(expectedPages, loaders.Keys.OrderBy(page => page).ToArray());
    }

    [Fact]
    public void AppShellModuleRegistrar_TrackedModuleStateLocksAndUnlocksNavigation()
    {
        var modules = CreateModules();

        AppShellModuleRegistrar.RegisterNavigationStateSources(modules);

        SetBooleanProperty(modules.ContactsViewModel, nameof(ContactsViewModel.IsLoading), true);
        Assert.True(modules.ShellViewModel.IsNavigationLocked);

        SetBooleanProperty(modules.ContactsViewModel, nameof(ContactsViewModel.IsLoading), false);
        Assert.False(modules.ShellViewModel.IsNavigationLocked);
    }

    [Fact]
    public void AppShellModuleRegistrar_RegistersExpectedUnsavedChangesChecks()
    {
        var modules = CreateModules();

        AppShellModuleRegistrar.RegisterUnsavedChangesChecks(modules);

        var navigationState = modules.ShellViewModel.Navigation;
        var checksField = typeof(ShellNavigationStateViewModel).GetField("_unsavedChangesChecks", BindingFlags.Instance | BindingFlags.NonPublic);

        Assert.NotNull(checksField);

        var checks = checksField!.GetValue(navigationState) as System.Collections.ICollection;
        Assert.NotNull(checks);
        Assert.Equal(4, checks!.Count);
    }

    private static AppModuleCatalog CreateModules()
    {
        var workerService = new BootstrapWorkerServiceStub();
        var navigationService = new NavigationService();
        var cacheService = new CacheService();
        var exchangeConfiguration = new ExchangeOnlineConfiguration();
        var shellViewModel = new ShellViewModel(workerService, navigationService, exchangeConfiguration);

        // Build a disabled catalog service (no network, no timer) for bootstrap tests.
        var catalogConfig = new LicenseCatalogConfiguration
        {
            AutoUpdateMode = CatalogAutoUpdateMode.Disabled,
            CheckOnStartup = false,
            LocalCachePath = Path.Combine(Path.GetTempPath(), "OnlyExo365.Tests.Bootstrap", Guid.NewGuid().ToString("N"))
        };
        var catalogFileStore = new LicenseCatalogFileStore(catalogConfig);
#pragma warning disable CA2000 // LicenseCatalogUpdateService owns and disposes the downloader.
        var catalogDownloader = new LicenseCatalogDownloader(catalogConfig);
        var skuNameResolver = new PresentationSkuNameResolver();
        LicenseCatalogUpdateService? catalogUpdateService = null;

        try
        {
            catalogUpdateService = new LicenseCatalogUpdateService(
                catalogConfig, catalogFileStore, catalogDownloader, skuNameResolver);
            var catalogViewModel = new LicenseCatalogViewModel(catalogUpdateService, shellViewModel);

            return new AppModuleCatalog
            {
                WorkerService = workerService,
                NavigationService = navigationService,
                LicenseCatalogUpdateService = catalogUpdateService,
                ShellViewModel = shellViewModel,
                LanguageSelectionViewModel = new LanguageSelectionViewModel(new UserPreferencesService()),
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
        catch
        {
            catalogUpdateService?.Dispose();
            catalogDownloader.Dispose();
            throw;
        }
#pragma warning restore CA2000
    }

    private static void SetBooleanProperty(object instance, string propertyName, bool value)
    {
        var property = instance.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

        Assert.NotNull(property);

        property!.SetValue(instance, value);
    }

    private sealed class BootstrapWorkerServiceStub : TestWorkerServiceBase;
}

