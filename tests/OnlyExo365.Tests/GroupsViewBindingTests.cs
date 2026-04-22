using System.Xml.Linq;

namespace OnlyExo365.Tests;

public sealed class GroupsViewBindingTests
{
    [Theory]
    [InlineData("DataContext.DistributionLists.RemoveMemberCommand")]
    [InlineData("DataContext.DistributionLists.RemoveAcceptSenderCommand")]
    [InlineData("DataContext.DistributionLists.RemoveRejectSenderCommand")]
    [InlineData("DistributionLists.PreviewDynamicMembersCommand")]
    public void DistributionListView_ExposesExpectedButtonCommands(string commandBinding)
    {
        var document = LoadViewDocument("DistributionListView.xaml");

        var button = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals(
                    (string?)candidate.Attribute("Command"),
                    commandBinding.StartsWith("DataContext.", StringComparison.Ordinal)
                        ? $"{{Binding {commandBinding}, RelativeSource={{RelativeSource AncestorType=UserControl}}}}"
                        : $"{{Binding {commandBinding}}}",
                    StringComparison.Ordinal));

        Assert.NotNull(button);
    }

    [Fact]
    public void DistributionListView_ShowsMembersStatusText()
    {
        var document = LoadViewDocument("DistributionListView.xaml");

        var textBlock = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "TextBlock" &&
                string.Equals((string?)candidate.Attribute("Text"), "{Binding DistributionLists.MembersStatusText}", StringComparison.Ordinal));

        Assert.NotNull(textBlock);
    }

    private static XDocument LoadViewDocument(string viewFileName)
    {
        var viewPath = TestPathHelper.GetRepositoryPath("src", "OnlyExo365.Shell", "Views", viewFileName);

        return XDocument.Load(viewPath);
    }
}

