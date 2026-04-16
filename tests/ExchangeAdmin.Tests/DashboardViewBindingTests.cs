using System.Xml.Linq;

namespace ExchangeAdmin.Tests;

public sealed class DashboardViewBindingTests
{
    [Fact]
    public void DashboardView_UsesVisibleNavigationButtonsForSummaryCards()
    {
        var document = LoadView("DashboardView.xaml");

        Assert.Contains(
            document.Descendants(),
            candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Content"), "{loc:Loc Key=Dashboard.UserMailboxes}", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Command"), "{Binding NavigateToMailboxesCommand}", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Style"), "{StaticResource SummaryLinkButton}", StringComparison.Ordinal));

        Assert.Contains(
            document.Descendants(),
            candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Content"), "{loc:Loc Key=Dashboard.SharedMailboxes}", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Command"), "{Binding NavigateToSharedMailboxesCommand}", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Style"), "{StaticResource SummaryLinkButton}", StringComparison.Ordinal));

        Assert.Contains(
            document.Descendants(),
            candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Content"), "{loc:Loc Key=Dashboard.DistributionGroups}", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Command"), "{Binding NavigateToDistributionListsCommand}", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Style"), "{StaticResource SummaryLinkButton}", StringComparison.Ordinal));

        Assert.DoesNotContain(
            document.Descendants(),
            candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Opacity"), "0", StringComparison.Ordinal) &&
                candidate.Attribute("Command")?.Value?.Contains("NavigateTo", StringComparison.Ordinal) == true);
    }

    private static XDocument LoadView(string fileName)
        => XDocument.Load(TestPathHelper.GetRepositoryPath("src", "ExchangeAdmin.Presentation", "Views", fileName));
}
