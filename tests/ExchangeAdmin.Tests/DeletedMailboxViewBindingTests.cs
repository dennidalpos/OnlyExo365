using System.Xml.Linq;

namespace ExchangeAdmin.Tests;

public sealed class DeletedMailboxViewBindingTests
{
    [Fact]
    public void DeletedMailboxesView_ShowsHardDeletedRecoveryHint()
    {
        var document = LoadViewDocument("DeletedMailboxesView.xaml");

        var hint = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "TextBlock" &&
                ((string?)candidate.Attribute("Text"))?.Contains("DeletedMailbox.ListDescription", StringComparison.Ordinal) == true);

        Assert.NotNull(hint);
    }

    [Fact]
    public void DeletedMailboxesView_ExposesVisibleLabelAndAssistiveMetadataForUpnLookup()
    {
        var document = LoadViewDocument("DeletedMailboxesView.xaml");

        var label = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "TextBlock" &&
                string.Equals((string?)candidate.Attribute("Text"), "{loc:Loc Key=DeletedMailbox.UpnLabel}", StringComparison.Ordinal));

        Assert.NotNull(label);

        var lookupBox = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "TextBox" &&
                string.Equals((string?)candidate.Attribute("Grid.Row"), "1", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Grid.Column"), "0", StringComparison.Ordinal));

        Assert.NotNull(lookupBox);

        var style = lookupBox!
            .Elements()
            .FirstOrDefault(candidate => candidate.Name.LocalName == "TextBox.Style")
            ?.Elements()
            .FirstOrDefault(candidate => candidate.Name.LocalName == "Style");

        Assert.NotNull(style);
        Assert.Contains(style!.Elements(), candidate =>
            candidate.Name.LocalName == "Setter" &&
            string.Equals((string?)candidate.Attribute("Property"), "AutomationProperties.Name", StringComparison.Ordinal) &&
            string.Equals((string?)candidate.Attribute("Value"), "{loc:Loc Key=DeletedMailbox.UpnLabel}", StringComparison.Ordinal));
        Assert.Contains(style.Elements(), candidate =>
            candidate.Name.LocalName == "Setter" &&
            string.Equals((string?)candidate.Attribute("Property"), "ToolTip", StringComparison.Ordinal) &&
            string.Equals((string?)candidate.Attribute("Value"), "{loc:Loc Key=DeletedMailbox.UpnTooltip}", StringComparison.Ordinal));

        var checkButton = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Command"), "{Binding DeletedMailboxes.CheckUpnCommand}", StringComparison.Ordinal));

        Assert.NotNull(checkButton);
        Assert.Equal("{loc:Loc Key=DeletedMailbox.CheckUpnTooltip}", (string?)checkButton!.Attribute("ToolTip"));
        Assert.Equal("{loc:Loc Key=DeletedMailbox.CheckUpn}", (string?)checkButton.Attribute("AutomationProperties.Name"));
    }

    [Fact]
    public void MailboxRestoreTab_ExplainsInactiveAndHardDeletedSemantics()
    {
        var document = LoadViewDocument("MailboxRestoreTab.xaml");

        var hint = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "TextBlock" &&
                ((string?)candidate.Attribute("Text"))?.Contains("MailboxRestore.Description", StringComparison.Ordinal) == true);

        Assert.NotNull(hint);
    }

    private static XDocument LoadViewDocument(string viewFileName)
    {
        var viewPath = TestPathHelper.GetRepositoryPath("src", "ExchangeAdmin.Presentation", "Views", viewFileName);

        return XDocument.Load(viewPath);
    }
}
