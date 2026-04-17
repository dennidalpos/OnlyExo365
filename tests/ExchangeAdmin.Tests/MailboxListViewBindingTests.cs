namespace ExchangeAdmin.Tests;

public sealed class MailboxListViewBindingTests
{
    [Fact]
    public void MailboxListView_UsesPasswordBoxForNewMailboxPassword()
    {
        var content = File.ReadAllText(GetViewPath());

        Assert.Contains("PasswordChanged=\"NewMailboxPasswordBox_OnPasswordChanged\"", content, StringComparison.Ordinal);
        Assert.Contains("helpers:PasswordBoxClearHelper.ClearTrigger=\"{Binding Mailboxes.NewMailboxPasswordClearTrigger}\"", content, StringComparison.Ordinal);
        Assert.DoesNotContain("Text=\"{Binding Mailboxes.NewMailboxPassword, UpdateSourceTrigger=PropertyChanged}\"", content, StringComparison.Ordinal);
    }

    [Fact]
    public void MailboxListView_ExposesProvisioningWorkspaceToggleAndLicensingPanel()
    {
        var content = File.ReadAllText(GetViewPath());

        Assert.Contains("{loc:Loc Key=Mailbox.List.UnlicensedMembersHeading}", content, StringComparison.Ordinal);
        Assert.Contains("{loc:Loc Key=Mailbox.List.UnlicensedMembersDescription}", content, StringComparison.Ordinal);
        Assert.Contains("DataContext=\"{Binding Mailboxes.Provisioning}\"", content, StringComparison.Ordinal);
    }

    [Fact]
    public void MailboxListView_ExposesDisconnectedHelperSearchAccessibilityAndDetailsEmptyState()
    {
        var content = File.ReadAllText(GetViewPath());

        Assert.Contains("{loc:Loc Key=Mailbox.List.InventoryDisconnectedMessage}", content, StringComparison.Ordinal);
        Assert.Contains("Property=\"AutomationProperties.Name\" Value=\"{loc:Loc Key=Mailbox.List.SearchLabel}\"", content, StringComparison.Ordinal);
        Assert.Contains("{loc:Loc Key=Mailbox.List.DetailsEmptyState}", content, StringComparison.Ordinal);
    }

    private static string GetViewPath()
    {
        return TestPathHelper.GetRepositoryPath("src", "ExchangeAdmin.Presentation", "Views", "MailboxListView.xaml");
    }
}
