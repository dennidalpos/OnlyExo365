using ExchangeAdmin.Presentation.Localization;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.Text;

public static class UiTextCatalog
{
    public static string ShellWindowTitle => Loc.Get("Shell.WindowTitle");
    public static string AppName => Loc.Get("Shell.AppName");
    public static string WorkerLabel => Loc.Get("Status.WorkerLabel");
    public static string ExchangeLabel => Loc.Get("Status.ExchangeLabel");
    public static string GraphLabel => Loc.Get("Status.GraphLabel");
    public static string DashboardNav => Loc.Get("Nav.Dashboard");
    public static string MailContactsNav => Loc.Get("Nav.MailContacts");
    public static string ResourceMailboxesNav => Loc.Get("Nav.ResourceMailboxes");
    public static string DistributionListsNav => Loc.Get("Nav.DistributionLists");
    public static string MailboxesNav => Loc.Get("Nav.Mailboxes");
    public static string DeletedMailboxesNav => Loc.Get("Nav.DeletedMailboxes");
    public static string MailboxStorageNav => Loc.Get("Nav.MailboxStorage");
    public static string MailboxAccessNav => Loc.Get("Nav.MailboxAccess");
    public static string PublicFoldersNav => Loc.Get("Nav.PublicFolders");
    public static string MobileDevicesNav => Loc.Get("Nav.MobileDevices");
    public static string MailFlowNav => Loc.Get("Nav.MailFlow");
    public static string MailSecurityNav => Loc.Get("Nav.MailSecurity");
    public static string MessageTraceNav => Loc.Get("Nav.MessageTrace");
    public static string MigrationNav => Loc.Get("Nav.Migration");
    public static string RoleGroupsNav => Loc.Get("Nav.RoleGroups");
    public static string ComplianceNav => Loc.Get("Nav.Compliance");
    public static string ToolsNav => Loc.Get("Nav.Tools");
    public static string LogsNav => Loc.Get("Nav.Logs");

    public static string StartWorkerButton => Loc.Get("Btn.StartWorker");
    public static string StopWorkerButton => Loc.Get("Btn.StopWorker");
    public static string ConnectExchangeButton => Loc.Get("Btn.ConnectExchange");
    public static string DisconnectButton => Loc.Get("Btn.Disconnect");

    public static string CancelButton => Loc.Get("Btn.Cancel");
    public static string ConfirmButton => Loc.Get("Btn.Confirm");
    public static string CloseButton => Loc.Get("Btn.Close");
    public static string OkButton => Loc.Get("Btn.Ok");
    public static string SaveButton => Loc.Get("Btn.Save");
    public static string RefreshButton => Loc.Get("Btn.Refresh");
    public static string LoadMoreButton => Loc.Get("Btn.LoadMore");
    public static string NewButton => Loc.Get("Btn.New");
    public static string RemoveButton => Loc.Get("Btn.Remove");
    public static string UpdateButton => Loc.Get("Btn.Update");

    public static string MailboxDetailsFallbackTitle => Loc.Get("Mailbox.DetailsFallbackTitle");
    public static string MailboxOverviewTab => Loc.Get("Tab.MailboxOverview");
    public static string MailboxSettingsTab => Loc.Get("Tab.MailboxSettings");
    public static string MailboxLicensesTab => Loc.Get("Tab.MailboxLicenses");
    public static string MailboxRestoreTab => Loc.Get("Tab.MailboxRestore");
    public static string MailboxPermissionsTab => Loc.Get("Tab.MailboxPermissions");
    public static string MailboxDetailsErrorTitle => Loc.Get("Mailbox.ErrorTitle");
    public static string MailboxDetailsLoadingTitle => Loc.Get("Mailbox.LoadingTitle");
    public static string MailboxDetailsLoadingOverlayTitle => Loc.Get("Mailbox.LoadingOverlayTitle");
    public static string MailboxDetailsSavingOverlayTitle => Loc.Get("Mailbox.SavingOverlayTitle");
    public static string BackToMailboxListButton => Loc.Get("Btn.BackToMailboxList");
    public static string BackToMailboxListTooltip => Loc.Get("Mailbox.BackTooltip");
    public static string MailboxRefreshTooltip => Loc.Get("Mailbox.RefreshTooltip");
    public static string PendingMailboxChangesLabel => Loc.Get("Mailbox.PendingChangesLabel");
    public static string PendingMailboxChangesTooltip => Loc.Get("Mailbox.PendingChangesTooltip");
    public static string PendingMailboxChangesDiscardTooltip => Loc.Get("Mailbox.PendingChangesDiscardTooltip");
    public static string PendingPermissionsTooltip => Loc.Get("Mailbox.PendingPermissionsTooltip");
    public static string PendingPermissionsDiscardTooltip => Loc.Get("Mailbox.PendingPermissionsDiscardTooltip");
    public static string PendingPermissionChangesFormat => Loc.Get("Mailbox.PendingPermissionsFormat");

    public static string GetNavigationLabel(NavigationPage page) => page switch
    {
        NavigationPage.Dashboard => Loc.Get("Nav.Dashboard"),
        NavigationPage.Contacts => Loc.Get("Nav.MailContacts"),
        NavigationPage.Resources => Loc.Get("Nav.ResourceMailboxes"),
        NavigationPage.PublicFolders => Loc.Get("Nav.PublicFolders"),
        NavigationPage.MobileDevices => Loc.Get("Nav.MobileDevices"),
        NavigationPage.Migration => Loc.Get("Nav.Migration"),
        NavigationPage.Permissions => Loc.Get("Nav.RoleGroups"),
        NavigationPage.Mailboxes => Loc.Get("Nav.Mailboxes"),
        NavigationPage.DeletedMailboxes => Loc.Get("Nav.DeletedMailboxes"),
        NavigationPage.MailboxSpace => Loc.Get("Nav.MailboxStorage"),
        NavigationPage.MailboxAccessReport => Loc.Get("Nav.MailboxAccess"),
        NavigationPage.DistributionLists => Loc.Get("Nav.DistributionLists"),
        NavigationPage.MessageTrace => Loc.Get("Nav.MessageTrace"),
        NavigationPage.Compliance => Loc.Get("Nav.Compliance"),
        NavigationPage.MailSecurity => Loc.Get("Nav.MailSecurity"),
        NavigationPage.MailFlow => Loc.Get("Nav.MailFlow"),
        NavigationPage.Tools => Loc.Get("Nav.Tools"),
        NavigationPage.Logs => Loc.Get("Nav.Logs"),
        _ => Loc.Get("Shell.AppName")
    };
}
