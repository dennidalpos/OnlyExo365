using OnlyExo365.Shell.Services;
using OnlyExo365.Shell.Ipc;
using OnlyExo365.Shell.ViewModels;

namespace OnlyExo365.Shell.Bootstrap;

internal sealed class AppModuleCatalog
{
    public required IWorkerService WorkerService { get; init; }
    public required NavigationService NavigationService { get; init; }
    public required LicenseCatalogUpdateService LicenseCatalogUpdateService { get; init; }
    public required ShellViewModel ShellViewModel { get; init; }
    public required LanguageSelectionViewModel LanguageSelectionViewModel { get; init; }
    public required DashboardViewModel DashboardViewModel { get; init; }
    public required ContactsViewModel ContactsViewModel { get; init; }
    public required ResourcesViewModel ResourcesViewModel { get; init; }
    public required PublicFoldersViewModel PublicFoldersViewModel { get; init; }
    public required MobileDevicesViewModel MobileDevicesViewModel { get; init; }
    public required MigrationViewModel MigrationViewModel { get; init; }
    public required PermissionsViewModel PermissionsViewModel { get; init; }
    public required MailboxListViewModel MailboxListViewModel { get; init; }
    public required DeletedMailboxesViewModel DeletedMailboxesViewModel { get; init; }
    public required MailboxDetailsViewModel MailboxDetailsViewModel { get; init; }
    public required MailboxSpaceViewModel MailboxSpaceViewModel { get; init; }
    public required MailboxAccessReportViewModel MailboxAccessReportViewModel { get; init; }
    public required DistributionListViewModel DistributionListViewModel { get; init; }
    public required LogsViewModel LogsViewModel { get; init; }
    public required ToolsViewModel ToolsViewModel { get; init; }
    public required MessageTraceViewModel MessageTraceViewModel { get; init; }
    public required ComplianceViewModel ComplianceViewModel { get; init; }
    public required MailSecurityViewModel MailSecurityViewModel { get; init; }
    public required MailFlowViewModel MailFlowViewModel { get; init; }
}

