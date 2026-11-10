using OnlyExo365.Shell.Services;

namespace OnlyExo365.Shell.Bootstrap;

internal static class AppPageLoaderCatalog
{
    internal static IReadOnlyDictionary<NavigationPage, Func<Task>> Create(AppModuleCatalog modules)
    {
        return new Dictionary<NavigationPage, Func<Task>>
        {
            [NavigationPage.Dashboard] = modules.DashboardViewModel.LoadAsync,
            [NavigationPage.Contacts] = modules.ContactsViewModel.LoadAsync,
            [NavigationPage.Resources] = modules.ResourcesViewModel.LoadAsync,
            [NavigationPage.PublicFolders] = modules.PublicFoldersViewModel.LoadAsync,
            [NavigationPage.MobileDevices] = modules.MobileDevicesViewModel.LoadAsync,
            [NavigationPage.Migration] = modules.MigrationViewModel.LoadAsync,
            [NavigationPage.Permissions] = modules.PermissionsViewModel.LoadAsync,
            [NavigationPage.Mailboxes] = modules.MailboxListViewModel.LoadAsync,
            [NavigationPage.DeletedMailboxes] = modules.DeletedMailboxesViewModel.LoadAsync,
            [NavigationPage.DistributionLists] = modules.DistributionListViewModel.LoadAsync,
            [NavigationPage.Tools] = modules.ToolsViewModel.LoadAsync,
            [NavigationPage.MessageTrace] = modules.MessageTraceViewModel.LoadAsync,
            [NavigationPage.Compliance] = modules.ComplianceViewModel.LoadAsync,
            [NavigationPage.MailSecurity] = modules.MailSecurityViewModel.LoadAsync,
            [NavigationPage.MailFlow] = modules.MailFlowViewModel.LoadAsync,
            [NavigationPage.Logs] = () =>
            {
                modules.LogsViewModel.Refresh();
                return Task.CompletedTask;
            }
        };
    }
}

