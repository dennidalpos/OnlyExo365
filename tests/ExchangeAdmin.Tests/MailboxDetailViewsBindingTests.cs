using System.Xml.Linq;

namespace ExchangeAdmin.Tests;

public sealed class MailboxDetailViewsBindingTests
{
    [Fact]
    public void MailboxDetailsView_ExposesPendingChangeBannersAndOperationalTabs()
    {
        var document = LoadView("MailboxDetailsView.xaml");

        Assert.Contains(
            document.Descendants().Where(element => element.Name.LocalName == "TextBlock").Select(element => (string?)element.Attribute("Text")),
            text => string.Equals(text, "{loc:Loc Key=Mailbox.PendingChangesLabel}", StringComparison.Ordinal));
        Assert.Contains(
            document.Descendants().Where(element => element.Name.LocalName == "TextBlock").Select(element => (string?)element.Attribute("Text")),
            text => string.Equals(text, "{Binding PendingActionsCount, StringFormat={loc:Loc Key=Mailbox.PendingPermissionsFormat}}", StringComparison.Ordinal));

        var tabHeaders = document
            .Descendants()
            .Where(element => element.Name.LocalName == "TabItem")
            .Select(element => element.Attribute("Header")?.Value)
            .Where(static header => !string.IsNullOrWhiteSpace(header))
            .Select(static header => header!)
            .ToArray();

        Assert.Equal(
            [
                "{loc:Loc Key=Tab.MailboxOverview}",
                "{loc:Loc Key=Tab.MailboxSettings}",
                "{loc:Loc Key=Tab.MailboxLicenses}",
                "{loc:Loc Key=Tab.MailboxRestore}",
                "{loc:Loc Key=Tab.MailboxPermissions}"
            ],
            tabHeaders);

        var viewReferences = document
            .Descendants()
            .Select(element => element.Name.LocalName)
            .Where(name => name is "MailboxOverviewTab" or "MailboxSettingsTab" or "MailboxLicensesTab" or "MailboxRestoreTab" or "MailboxPermissionsTab")
            .ToArray();

        Assert.Equal(
            ["MailboxOverviewTab", "MailboxSettingsTab", "MailboxLicensesTab", "MailboxRestoreTab", "MailboxPermissionsTab"],
            viewReferences);
    }

    [Fact]
    public void MailboxOverviewTab_ExposesConversionActionsAndForwardingWarnings()
    {
        var content = File.ReadAllText(GetViewPath("MailboxOverviewTab.xaml"));

        Assert.Contains("{loc:Loc Key=Mailbox.Overview.ConvertToShared}", content, StringComparison.Ordinal);
        Assert.Contains("{loc:Loc Key=Mailbox.Overview.RestoreUserMailbox}", content, StringComparison.Ordinal);
        Assert.Contains("{loc:Loc Key=Mailbox.Overview.ForwardingRuleDetected}", content, StringComparison.Ordinal);
        Assert.Contains("{loc:Loc Key=Mailbox.Overview.RedirectRuleDetected}", content, StringComparison.Ordinal);
        Assert.Contains("Visibility=\"{Binding HasConversionActions, Converter={StaticResource BoolToVisibility}}\"", content, StringComparison.Ordinal);
    }

    [Fact]
    public void MailboxPermissionsTab_ExposesAutoMappingAndFolderPermissionWorkflows()
    {
        var content = File.ReadAllText(GetViewPath("MailboxPermissionsTab.xaml"));

        Assert.Contains("SelectedValue=\"{Binding NewPermissionAutoMapping}\"", content, StringComparison.Ordinal);
        Assert.Contains("ModifyAutoMappingCommand", content, StringComparison.Ordinal);
        Assert.Contains("RefreshFolderPermissionsCommand", content, StringComparison.Ordinal);
        Assert.Contains("FolderPermissionFolderPath", content, StringComparison.Ordinal);
        Assert.Contains("AddFolderPermissionCommand", content, StringComparison.Ordinal);
        Assert.Contains("UpdateFolderPermissionCommand", content, StringComparison.Ordinal);
        Assert.Contains("RemoveFolderPermissionCommand", content, StringComparison.Ordinal);
    }

    [Fact]
    public void MailboxPermissionsTab_UsesEnglishPermissionLabels()
    {
        var content = File.ReadAllText(GetViewPath("MailboxPermissionsTab.xaml"));

        Assert.Contains("{loc:Loc Key=MailboxPerms.ManagementHeading}", content, StringComparison.Ordinal);
        Assert.Contains("{loc:Loc Key=MailboxPerms.PermissionFullAccess}", content, StringComparison.Ordinal);
        Assert.Contains("{loc:Loc Key=MailboxPerms.PermissionSendAs}", content, StringComparison.Ordinal);
        Assert.Contains("{loc:Loc Key=MailboxPerms.PermissionSendOnBehalf}", content, StringComparison.Ordinal);
        Assert.Contains("{loc:Loc Key=MailboxPerms.ColInherited}", content, StringComparison.Ordinal);
        Assert.DoesNotContain("autorizz", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Accesso", content, StringComparison.Ordinal);
        Assert.DoesNotContain("Invia", content, StringComparison.Ordinal);
        Assert.DoesNotContain("Ereditato", content, StringComparison.Ordinal);
    }

    [Fact]
    public void MailboxSettingsTab_ExposesRetentionCasAndAutoReplyEditors()
    {
        var content = File.ReadAllText(GetViewPath("MailboxSettingsTab.xaml"));

        Assert.Contains("SettingsEditor.ProxyAddressesText", content, StringComparison.Ordinal);
        Assert.Contains("SettingsEditor.CasCapabilityMessage", content, StringComparison.Ordinal);
        Assert.Contains("RefreshRetentionPoliciesCommand", content, StringComparison.Ordinal);
        Assert.Contains("ShowArchiveRequiredWarning", content, StringComparison.Ordinal);
        Assert.Contains("AutoReplyEditor.AutoReplyEnabled", content, StringComparison.Ordinal);
        Assert.Contains("AutoReplyEditor.AutoReplyExternalAudience", content, StringComparison.Ordinal);
    }

    [Fact]
    public void ToolsView_ExposesLeastPrivilegeMatrixAndModuleBootstrapActions()
    {
        var content = File.ReadAllText(GetViewPath("ToolsView.xaml"));

        Assert.Contains("{loc:Loc Key=Tools.LeastPrivilegeMatrix}", content, StringComparison.Ordinal);
        Assert.Contains("ItemsSource=\"{Binding Tools.LeastPrivilegeMatrix}\"", content, StringComparison.Ordinal);
        Assert.Contains("{loc:Loc Key=Tools.ValidationPrefix}", content, StringComparison.Ordinal);
        Assert.Contains("{loc:Loc Key=Tools.CheckPrerequisites}", content, StringComparison.Ordinal);
        Assert.Contains("InstallPowerShellCommand", content, StringComparison.Ordinal);
        Assert.Contains("InstallExchangeModuleCommand", content, StringComparison.Ordinal);
        Assert.Contains("InstallGraphModuleCommand", content, StringComparison.Ordinal);
        Assert.Contains("{loc:Loc Key=Tools.WorkerConsoleToggle}", content, StringComparison.Ordinal);
        Assert.Contains("Checked=\"WorkerConsoleToggle_Changed\"", content, StringComparison.Ordinal);
        Assert.Contains("Unchecked=\"WorkerConsoleToggle_Changed\"", content, StringComparison.Ordinal);
        Assert.Contains("IsWorkerConsoleVisible", content, StringComparison.Ordinal);
        Assert.DoesNotContain("CommandParameter=\"{Binding RelativeSource={RelativeSource Self}, Path=IsChecked}\"", content, StringComparison.Ordinal);
        Assert.Contains("{loc:Loc Key=Tools.ExchangeOnlineConnectionsAreDisabledByPolicyEXCHANGEADMINDISABLE}", content, StringComparison.Ordinal);
        Assert.DoesNotContain("Quick Actions", content, StringComparison.Ordinal);
        Assert.DoesNotContain("Width=\"760\"", content, StringComparison.Ordinal);
    }

    private static XDocument LoadView(string fileName)
        => XDocument.Load(GetViewPath(fileName));

    private static string GetViewPath(string fileName)
        => TestPathHelper.GetRepositoryPath("src", "ExchangeAdmin.Presentation", "Views", fileName);
}
