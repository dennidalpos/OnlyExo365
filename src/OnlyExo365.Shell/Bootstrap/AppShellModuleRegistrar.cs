using OnlyExo365.Shell.ViewModels;

namespace OnlyExo365.Shell.Bootstrap;

internal static class AppShellModuleRegistrar
{
    public static void Configure(AppModuleCatalog modules)
    {
        AssignModules(modules);
        RegisterNavigationStateSources(modules);
        RegisterUnsavedChangesChecks(modules);
        WireLicenseCatalogRelay(modules);
    }

    internal static void WireLicenseCatalogRelay(AppModuleCatalog modules)
    {
        modules.LicenseCatalogUpdateService.CatalogUpdated += (_, args) =>
            modules.ShellViewModel.RaiseLicenseCatalogUpdated(args);
    }

    internal static void AssignModules(AppModuleCatalog modules)
    {
        var shellViewModel = modules.ShellViewModel;

        shellViewModel.LanguageSelection = modules.LanguageSelectionViewModel;
        shellViewModel.Dashboard = modules.DashboardViewModel;
        shellViewModel.Contacts = modules.ContactsViewModel;
        shellViewModel.Resources = modules.ResourcesViewModel;
        shellViewModel.PublicFolders = modules.PublicFoldersViewModel;
        shellViewModel.MobileDevices = modules.MobileDevicesViewModel;
        shellViewModel.Migration = modules.MigrationViewModel;
        shellViewModel.Permissions = modules.PermissionsViewModel;
        shellViewModel.Mailboxes = modules.MailboxListViewModel;
        shellViewModel.DeletedMailboxes = modules.DeletedMailboxesViewModel;
        shellViewModel.MailboxDetails = modules.MailboxDetailsViewModel;
        shellViewModel.MailboxSpace = modules.MailboxSpaceViewModel;
        shellViewModel.MailboxAccessReport = modules.MailboxAccessReportViewModel;
        shellViewModel.DistributionLists = modules.DistributionListViewModel;
        shellViewModel.Logs = modules.LogsViewModel;
        shellViewModel.Tools = modules.ToolsViewModel;
        shellViewModel.MessageTrace = modules.MessageTraceViewModel;
        shellViewModel.Compliance = modules.ComplianceViewModel;
        shellViewModel.MailSecurity = modules.MailSecurityViewModel;
        shellViewModel.MailFlow = modules.MailFlowViewModel;
    }

    internal static void RegisterNavigationStateSources(AppModuleCatalog modules)
    {
        var shellViewModel = modules.ShellViewModel;

        shellViewModel.RegisterNavigationStateSource(
            modules.DashboardViewModel,
            () => modules.DashboardViewModel.IsLoading,
            nameof(DashboardViewModel.IsLoading));
        shellViewModel.RegisterNavigationStateSource(
            modules.ContactsViewModel,
            () => modules.ContactsViewModel.IsLoading || modules.ContactsViewModel.IsSaving,
            nameof(ContactsViewModel.IsLoading),
            nameof(ContactsViewModel.IsSaving));
        shellViewModel.RegisterNavigationStateSource(
            modules.ResourcesViewModel,
            () => modules.ResourcesViewModel.IsLoading || modules.ResourcesViewModel.IsSaving,
            nameof(ResourcesViewModel.IsLoading),
            nameof(ResourcesViewModel.IsSaving));
        shellViewModel.RegisterNavigationStateSource(
            modules.PublicFoldersViewModel,
            () => modules.PublicFoldersViewModel.IsLoading || modules.PublicFoldersViewModel.IsSaving,
            nameof(PublicFoldersViewModel.IsLoading),
            nameof(PublicFoldersViewModel.IsSaving));
        shellViewModel.RegisterNavigationStateSource(
            modules.MobileDevicesViewModel,
            () => modules.MobileDevicesViewModel.IsLoading || modules.MobileDevicesViewModel.IsApplyingAction,
            nameof(MobileDevicesViewModel.IsLoading),
            nameof(MobileDevicesViewModel.IsApplyingAction));
        shellViewModel.RegisterNavigationStateSource(
            modules.MigrationViewModel,
            () => modules.MigrationViewModel.IsLoading ||
                  modules.MigrationViewModel.IsLoadingDetails ||
                  modules.MigrationViewModel.IsApplyingAction ||
                  modules.MigrationViewModel.IsLoadingEndpoints ||
                  modules.MigrationViewModel.IsSavingEndpoint ||
                  modules.MigrationViewModel.IsTestingEndpoint ||
                  modules.MigrationViewModel.IsRunningPreflight ||
                  modules.MigrationViewModel.IsCreatingBatch,
            nameof(MigrationViewModel.IsLoading),
            nameof(MigrationViewModel.IsLoadingDetails),
            nameof(MigrationViewModel.IsApplyingAction),
            nameof(MigrationViewModel.IsLoadingEndpoints),
            nameof(MigrationViewModel.IsSavingEndpoint),
            nameof(MigrationViewModel.IsTestingEndpoint),
            nameof(MigrationViewModel.IsRunningPreflight),
            nameof(MigrationViewModel.IsCreatingBatch));
        shellViewModel.RegisterNavigationStateSource(
            modules.PermissionsViewModel,
            () => modules.PermissionsViewModel.IsLoading || modules.PermissionsViewModel.IsLoadingDetails || modules.PermissionsViewModel.IsSaving,
            nameof(PermissionsViewModel.IsLoading),
            nameof(PermissionsViewModel.IsLoadingDetails),
            nameof(PermissionsViewModel.IsSaving));
        shellViewModel.RegisterNavigationStateSource(
            modules.MailboxListViewModel,
            () => modules.MailboxListViewModel.IsLoading ||
                  modules.MailboxListViewModel.IsProvisioningLoading ||
                  modules.MailboxListViewModel.IsAssigningProvisioningLicense,
            nameof(MailboxListViewModel.IsLoading),
            nameof(MailboxListViewModel.IsProvisioningLoading),
            nameof(MailboxListViewModel.IsAssigningProvisioningLicense));
        shellViewModel.RegisterNavigationStateSource(
            modules.DeletedMailboxesViewModel,
            () => modules.DeletedMailboxesViewModel.IsLoading,
            nameof(DeletedMailboxesViewModel.IsLoading));
        shellViewModel.RegisterNavigationStateSource(
            modules.MailboxDetailsViewModel,
            () => modules.MailboxDetailsViewModel.IsLoading || modules.MailboxDetailsViewModel.IsRetentionPolicyLoading || modules.MailboxDetailsViewModel.IsSaving,
            nameof(MailboxDetailsViewModel.IsLoading),
            nameof(MailboxDetailsViewModel.IsRetentionPolicyLoading),
            nameof(MailboxDetailsViewModel.IsSaving));
        shellViewModel.RegisterNavigationStateSource(
            modules.MailboxSpaceViewModel,
            () => modules.MailboxSpaceViewModel.IsLoading,
            nameof(MailboxSpaceViewModel.IsLoading));
        shellViewModel.RegisterNavigationStateSource(
            modules.DistributionListViewModel,
            () => modules.DistributionListViewModel.IsLoading || modules.DistributionListViewModel.IsLoadingDetails || modules.DistributionListViewModel.IsLoadingMembers,
            nameof(DistributionListViewModel.IsLoading),
            nameof(DistributionListViewModel.IsLoadingDetails),
            nameof(DistributionListViewModel.IsLoadingMembers));
        shellViewModel.RegisterNavigationStateSource(
            modules.ComplianceViewModel,
            () => modules.ComplianceViewModel.IsBusy,
            nameof(ComplianceViewModel.IsLoadingWorkspace),
            nameof(ComplianceViewModel.IsSearchingAudit),
            nameof(ComplianceViewModel.IsCreatingSearch),
            nameof(ComplianceViewModel.IsApplyingAction));
        shellViewModel.RegisterNavigationStateSource(
            modules.MailSecurityViewModel,
            () => modules.MailSecurityViewModel.IsBusy,
            nameof(MailSecurityViewModel.IsLoadingWorkspace),
            nameof(MailSecurityViewModel.IsSaving));
    }

    internal static void RegisterUnsavedChangesChecks(AppModuleCatalog modules)
    {
        var shellViewModel = modules.ShellViewModel;

        shellViewModel.RegisterUnsavedChangesCheck(() => modules.MailboxDetailsViewModel.HasPendingChanges || modules.MailboxDetailsViewModel.HasPendingMailboxChanges);
        shellViewModel.RegisterUnsavedChangesCheck(() => modules.DistributionListViewModel.HasPendingSettingsChanges);
        shellViewModel.RegisterUnsavedChangesCheck(() => modules.ResourcesViewModel.HasPendingChanges);
        shellViewModel.RegisterUnsavedChangesCheck(() => modules.PublicFoldersViewModel.HasPendingChanges);
    }
}

