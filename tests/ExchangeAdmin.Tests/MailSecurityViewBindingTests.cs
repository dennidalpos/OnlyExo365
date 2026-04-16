using System.Xml.Linq;

namespace ExchangeAdmin.Tests;

public sealed class MailSecurityViewBindingTests
{
    [Fact]
    public void MailSecurityView_ExposesRefreshAndOutboundSaveCommands()
    {
        var document = LoadViewDocument();

        var refreshButton = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Content"), "{loc:Loc Key=MailSecurity.RefreshWorkspaceBtn}", StringComparison.Ordinal));

        Assert.NotNull(refreshButton);
        Assert.Equal("{Binding MailSecurity.RefreshWorkspaceCommand}", (string?)refreshButton!.Attribute("Command"));

        var saveButton = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Content"), "{loc:Loc Key=MailSecurity.SaveOutboundSpamPolicy}", StringComparison.Ordinal));

        Assert.NotNull(saveButton);
        Assert.Equal("{Binding MailSecurity.SaveOutboundSpamCommand}", (string?)saveButton!.Attribute("Command"));
    }

    private static XDocument LoadViewDocument()
    {
        var viewPath = TestPathHelper.GetRepositoryPath("src", "ExchangeAdmin.Presentation", "Views", "MailSecurityView.xaml");

        return XDocument.Load(viewPath);
    }
}
