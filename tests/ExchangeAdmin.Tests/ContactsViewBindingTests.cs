namespace ExchangeAdmin.Tests;

public sealed class ContactsViewBindingTests
{
    [Fact]
    public void ContactsView_UsesPasswordBoxClearTriggerForMailUserPassword()
    {
        var content = File.ReadAllText(GetViewPath());

        Assert.Contains("PasswordChanged=\"MailUserPasswordBox_OnPasswordChanged\"", content, StringComparison.Ordinal);
        Assert.Contains("helpers:PasswordBoxClearHelper.ClearTrigger=\"{Binding Contacts.MailUserPasswordClearTrigger}\"", content, StringComparison.Ordinal);
        Assert.Contains("Binding Contacts.IsCreateMailUser", content, StringComparison.Ordinal);
        Assert.DoesNotContain("Text=\"{Binding Contacts.Password, UpdateSourceTrigger=PropertyChanged}\"", content, StringComparison.Ordinal);
    }

    [Fact]
    public void ContactsView_DoesNotExposeDuplicateNewButtonInDetailsPanel()
    {
        var content = File.ReadAllText(GetViewPath());
        const string commandBinding = "Command=\"{Binding Contacts.NewContactCommand}\"";

        Assert.Equal(1, content.Split(commandBinding, StringSplitOptions.None).Length - 1);
    }

    private static string GetViewPath()
    {
        return TestPathHelper.GetRepositoryPath("src", "ExchangeAdmin.Presentation", "Views", "ContactsView.xaml");
    }
}
