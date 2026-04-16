using System.Xml.Linq;

namespace ExchangeAdmin.Tests;

public sealed class MailFlowViewBindingTests
{
    [Theory]
    [InlineData("MailFlow.NewConnectorCommand")]
    [InlineData("MailFlow.NewDomainCommand")]
    [InlineData("MailFlow.NewRemoteDomainCommand")]
    [InlineData("MailFlow.NewOrganizationRelationshipCommand")]
    [InlineData("MailFlow.NewAddressListCommand")]
    [InlineData("MailFlow.NewAddressBookPolicyCommand")]
    [InlineData("MailFlow.NewOfflineAddressBookCommand")]
    [InlineData("MailFlow.NewSharingPolicyCommand")]
    public void MailFlowView_ExposesResetCommandsForCreatePaths(string commandBinding)
    {
        var document = LoadViewDocument();

        var button = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Content"), "{loc:Loc Key=MailFlow.NewReset}", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Command"), $"{{Binding {commandBinding}}}", StringComparison.Ordinal));

        Assert.NotNull(button);
    }

    private static XDocument LoadViewDocument()
    {
        var viewPath = TestPathHelper.GetRepositoryPath("src", "ExchangeAdmin.Presentation", "Views", "MailFlowView.xaml");

        return XDocument.Load(viewPath);
    }
}
