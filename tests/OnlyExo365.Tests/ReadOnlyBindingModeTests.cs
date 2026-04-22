using System.Xml.Linq;

namespace OnlyExo365.Tests;

public sealed class ReadOnlyBindingModeTests
{
    [Theory]
    [InlineData("PublicFoldersView.xaml", "CheckBox", "Content", "{loc:Loc Key=PublicFolders.HasSubfolders}", "IsChecked", "{Binding PublicFolders.HasSubFolders, Mode=OneWay}")]
    [InlineData("PublicFoldersView.xaml", "TextBox", "Text", "{Binding PublicFolders.PermissionsSummary, Mode=OneWay}", "IsReadOnly", "True")]
    [InlineData("ComplianceView.xaml", "TextBlock", "Text", "{Binding Compliance.SelectedSearchExchangeLocations, Mode=OneWay}", "Foreground", "{StaticResource SecondaryTextBrush}")]
    [InlineData("ComplianceView.xaml", "TextBlock", "Text", "{Binding Compliance.SelectedSearchQuery, Mode=OneWay}", "TextWrapping", "Wrap")]
    [InlineData("MigrationView.xaml", "TextBox", "Text", "{Binding Migration.NotificationEmailsText, Mode=OneWay}", "IsReadOnly", "True")]
    [InlineData("MigrationView.xaml", "TextBox", "Text", "{Binding Migration.EndpointTestSummary, Mode=OneWay}", "IsReadOnly", "True")]
    [InlineData("MigrationView.xaml", "TextBox", "Text", "{Binding Migration.BatchPreflightSummary, Mode=OneWay}", "IsReadOnly", "True")]
    [InlineData("PermissionsView.xaml", "TextBox", "Text", "{Binding Permissions.RolesText, Mode=OneWay}", "IsReadOnly", "True")]
    [InlineData("PermissionsView.xaml", "TextBox", "Text", "{Binding Permissions.ScopesText, Mode=OneWay}", "IsReadOnly", "True")]
    [InlineData("PermissionsView.xaml", "TextBlock", "Text", "{Binding Permissions.SelectedRoleGroupDetails.ManagedBy.Count, Mode=OneWay}", "Text", "{Binding Permissions.SelectedRoleGroupDetails.ManagedBy.Count, Mode=OneWay}")]
    public void ReadOnlyViewBindings_UseOneWayMode(
        string viewFileName,
        string elementName,
        string keyAttributeName,
        string keyAttributeValue,
        string requiredAttributeName,
        string requiredAttributeValue)
    {
        var document = LoadViewDocument(viewFileName);

        var element = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == elementName &&
                string.Equals((string?)candidate.Attribute(keyAttributeName), keyAttributeValue, StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute(requiredAttributeName), requiredAttributeValue, StringComparison.Ordinal));

        Assert.NotNull(element);
    }

    private static XDocument LoadViewDocument(string viewFileName)
    {
        var viewPath = TestPathHelper.GetRepositoryPath("src", "OnlyExo365.Shell", "Views", viewFileName);

        return XDocument.Load(viewPath);
    }
}

