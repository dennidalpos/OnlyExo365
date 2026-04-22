using System.Xml.Linq;

namespace OnlyExo365.Tests;

public sealed class PublicFoldersViewBindingTests
{
    [Fact]
    public void PublicFoldersHasSubFoldersIndicator_UsesOneWayBinding()
    {
        var document = LoadViewDocument();

        var checkBox = document
            .Descendants()
            .FirstOrDefault(element =>
                element.Name.LocalName == "CheckBox" &&
                string.Equals((string?)element.Attribute("Content"), "{loc:Loc Key=PublicFolders.HasSubfolders}", StringComparison.Ordinal));

        Assert.NotNull(checkBox);

        var binding = (string?)checkBox!.Attribute("IsChecked");
        Assert.NotNull(binding);
        Assert.Contains("PublicFolders.HasSubFolders", binding, StringComparison.Ordinal);
        Assert.Contains("Mode=OneWay", binding, StringComparison.Ordinal);
    }

    [Fact]
    public void PublicFoldersPermissionsSummaryText_UsesOneWayBinding()
    {
        var document = LoadViewDocument();

        var textBox = document
            .Descendants()
            .FirstOrDefault(element =>
                element.Name.LocalName == "TextBox" &&
                string.Equals((string?)element.Attribute("IsReadOnly"), "True", StringComparison.Ordinal) &&
                string.Equals((string?)element.Attribute("Text"), "{Binding PublicFolders.PermissionsSummary, Mode=OneWay}", StringComparison.Ordinal));

        Assert.NotNull(textBox);
    }

    private static XDocument LoadViewDocument()
    {
        var viewPath = TestPathHelper.GetRepositoryPath("src", "OnlyExo365.Shell", "Views", "PublicFoldersView.xaml");
        return XDocument.Load(viewPath);
    }
}

