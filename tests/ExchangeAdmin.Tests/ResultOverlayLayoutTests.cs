using System.Xml.Linq;

namespace ExchangeAdmin.Tests;

public sealed class ResultOverlayLayoutTests
{
    [Theory]
    [InlineData("DashboardView.xaml", "{Binding Dashboard.IsLoading, Converter={StaticResource BoolToVisibility}}")]
    [InlineData("DistributionListView.xaml", "{Binding DistributionLists.IsBusy, Converter={StaticResource BoolToVisibility}}")]
    [InlineData("MessageTraceView.xaml", "{Binding MessageTrace.IsLoading, Converter={StaticResource BoolToVisibility}}")]
    [InlineData("MailboxSpaceView.xaml", "{Binding MailboxSpace.IsLoading, Converter={StaticResource BoolToVisibility}}")]
    [InlineData("MailboxAccessReportView.xaml", "{Binding MailboxAccessReport.IsLoading, Converter={StaticResource BoolToVisibility}}")]
    [InlineData("MobileDevicesView.xaml", "{Binding MobileDevices.IsBusy, Converter={StaticResource BoolToVisibility}}")]
    [InlineData("ComplianceView.xaml", "{Binding Compliance.IsLoadingOverlayVisible, Converter={StaticResource BoolToVisibility}}")]
    [InlineData("ContactsView.xaml", "{Binding Contacts.IsLoading, Converter={StaticResource BoolToVisibility}}")]
    [InlineData("MailSecurityView.xaml", "{Binding MailSecurity.IsBusy, Converter={StaticResource BoolToVisibility}}")]
    [InlineData("MailFlowView.xaml", "{Binding MailFlow.IsLoading, Converter={StaticResource BoolToVisibility}}")]
    [InlineData("PermissionsView.xaml", "{Binding Permissions.IsLoading, Converter={StaticResource BoolToVisibility}}")]
    public void ResultViews_UseCompactLoadingOverlayCard(string viewFileName, string visibilityBinding)
    {
        var document = LoadViewDocument(viewFileName);

        var overlayContainer = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName is "Border" or "Grid" &&
                string.Equals((string?)candidate.Attribute("Visibility"), visibilityBinding, StringComparison.Ordinal));

        Assert.NotNull(overlayContainer);

        var compactCard = overlayContainer!
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Border" &&
                (string.Equals((string?)candidate.Attribute("Style"), "{StaticResource CompactLoadingOverlayCardStyle}", StringComparison.Ordinal) ||
                 (string.Equals((string?)candidate.Attribute("Padding"), "14", StringComparison.Ordinal) &&
                  string.Equals((string?)candidate.Attribute("MinWidth"), "240", StringComparison.Ordinal) &&
                  string.Equals((string?)candidate.Attribute("MaxWidth"), "320", StringComparison.Ordinal))));

        Assert.NotNull(compactCard);

        var progressBar = compactCard!
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "ProgressBar" &&
                (string.Equals((string?)candidate.Attribute("Style"), "{StaticResource CompactLoadingOverlayProgressBarStyle}", StringComparison.Ordinal) ||
                 string.Equals((string?)candidate.Attribute("Width"), "220", StringComparison.Ordinal)));

        Assert.NotNull(progressBar);
    }

    [Fact]
    public void DarkTheme_DefinesSharedCompactLoadingOverlayStyles()
    {
        var themePath = TestPathHelper.GetRepositoryPath("src", "ExchangeAdmin.Presentation", "Themes", "DarkTheme.xaml");

        var document = XDocument.Load(themePath);

        Assert.Contains(document.Descendants(), element => element.Name.LocalName == "Style" && (string?)element.Attribute(XName.Get("Key", "http://schemas.microsoft.com/winfx/2006/xaml")) == "LoadingOverlayBackdropBorderStyle");
        Assert.Contains(document.Descendants(), element => element.Name.LocalName == "Style" && (string?)element.Attribute(XName.Get("Key", "http://schemas.microsoft.com/winfx/2006/xaml")) == "CompactLoadingOverlayCardStyle");
        Assert.Contains(document.Descendants(), element => element.Name.LocalName == "Style" && (string?)element.Attribute(XName.Get("Key", "http://schemas.microsoft.com/winfx/2006/xaml")) == "CompactLoadingOverlayProgressBarStyle");
        Assert.Contains(document.Descendants(), element => element.Name.LocalName == "Style" && (string?)element.Attribute(XName.Get("Key", "http://schemas.microsoft.com/winfx/2006/xaml")) == "CompactLoadingOverlayMetricsPanelStyle");
        Assert.Contains(document.Descendants(), element => element.Name.LocalName == "Style" && (string?)element.Attribute(XName.Get("Key", "http://schemas.microsoft.com/winfx/2006/xaml")) == "CompactLoadingOverlayCountTextStyle");
    }

    [Theory]
    [InlineData("DistributionListView.xaml", "{Binding DistributionLists.IsBusy, Converter={StaticResource BoolToVisibility}}")]
    [InlineData("MailboxSpaceView.xaml", "{Binding MailboxSpace.IsLoading, Converter={StaticResource BoolToVisibility}}")]
    [InlineData("MailboxAccessReportView.xaml", "{Binding MailboxAccessReport.IsLoading, Converter={StaticResource BoolToVisibility}}")]
    public void PrimaryResultViews_UseSharedCompactLoadingOverlayStyles(string viewFileName, string visibilityBinding)
    {
        var document = LoadViewDocument(viewFileName);

        var overlayContainer = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Border" &&
                string.Equals((string?)candidate.Attribute("Visibility"), visibilityBinding, StringComparison.Ordinal));

        Assert.NotNull(overlayContainer);
        Assert.Equal("{StaticResource LoadingOverlayBackdropBorderStyle}", (string?)overlayContainer!.Attribute("Style"));

        var compactCard = overlayContainer
            .Elements()
            .FirstOrDefault(candidate => candidate.Name.LocalName == "Border");

        Assert.NotNull(compactCard);
        Assert.Equal("{StaticResource CompactLoadingOverlayCardStyle}", (string?)compactCard!.Attribute("Style"));

        var statusText = compactCard
            .Descendants()
            .FirstOrDefault(candidate => candidate.Name.LocalName == "TextBlock" && candidate.Attribute("Text") != null);

        Assert.NotNull(statusText);
        Assert.Equal("{StaticResource CompactLoadingOverlayStatusTextStyle}", (string?)statusText!.Attribute("Style"));

        var progressBar = compactCard
            .Descendants()
            .FirstOrDefault(candidate => candidate.Name.LocalName == "ProgressBar");

        Assert.NotNull(progressBar);
        Assert.Equal("{StaticResource CompactLoadingOverlayProgressBarStyle}", (string?)progressBar!.Attribute("Style"));
    }

    [Fact]
    public void DashboardView_LoadingOverlay_SpansResultRowsWithoutDedicatedLayoutRow()
    {
        var document = LoadViewDocument("DashboardView.xaml");

        var overlayGrid = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Grid" &&
                string.Equals((string?)candidate.Attribute("Grid.Row"), "2", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Grid.RowSpan"), "3", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Visibility"), "{Binding Dashboard.IsLoading, Converter={StaticResource BoolToVisibility}}", StringComparison.Ordinal));

        Assert.NotNull(overlayGrid);
    }

    [Fact]
    public void MessageTraceView_LoadingOverlay_IsInsideResultsGrid()
    {
        var document = LoadViewDocument("MessageTraceView.xaml");

        var resultsGrid = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Grid" &&
                string.Equals((string?)candidate.Attribute("Grid.Row"), "5", StringComparison.Ordinal) &&
                candidate.Elements().Any(child =>
                    child.Name.LocalName == "Border" &&
                    string.Equals((string?)child.Attribute("Visibility"), "{Binding MessageTrace.IsLoading, Converter={StaticResource BoolToVisibility}}", StringComparison.Ordinal)));

        Assert.NotNull(resultsGrid);
    }

    [Fact]
    public void MobileDevicesView_LoadingOverlay_IsInsideDeviceListGrid()
    {
        var document = LoadViewDocument("MobileDevicesView.xaml");

        var resultsGrid = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Grid" &&
                string.Equals((string?)candidate.Attribute("Grid.Row"), "3", StringComparison.Ordinal) &&
                candidate.Elements().Any(child =>
                    child.Name.LocalName == "Border" &&
                    string.Equals((string?)child.Attribute("Visibility"), "{Binding MobileDevices.IsBusy, Converter={StaticResource BoolToVisibility}}", StringComparison.Ordinal)));

        Assert.NotNull(resultsGrid);
    }

    [Fact]
    public void ComplianceView_LoadingOverlay_StaysInsideWorkspaceRow()
    {
        var document = LoadViewDocument("ComplianceView.xaml");

        var overlay = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Border" &&
                string.Equals((string?)candidate.Attribute("Grid.Row"), "1", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Visibility"), "{Binding Compliance.IsLoadingOverlayVisible, Converter={StaticResource BoolToVisibility}}", StringComparison.Ordinal));

        Assert.NotNull(overlay);
    }

    [Fact]
    public void MailFlowView_LoadingOverlay_SpansBothWorkspaceColumns()
    {
        var document = LoadViewDocument("MailFlowView.xaml");

        var overlay = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Border" &&
                string.Equals((string?)candidate.Attribute("Grid.ColumnSpan"), "3", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Visibility"), "{Binding MailFlow.IsLoading, Converter={StaticResource BoolToVisibility}}", StringComparison.Ordinal));

        Assert.NotNull(overlay);
    }

    [Fact]
    public void ContactsView_LoadingOverlay_IsInsideContactsResultsGrid()
    {
        var document = LoadViewDocument("ContactsView.xaml");

        var resultsGrid = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Grid" &&
                string.Equals((string?)candidate.Attribute("Grid.Row"), "3", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Grid.RowSpan"), "2", StringComparison.Ordinal) &&
                candidate.Elements().Any(child =>
                    child.Name.LocalName == "Border" &&
                    string.Equals((string?)child.Attribute("Visibility"), "{Binding Contacts.IsLoading, Converter={StaticResource BoolToVisibility}}", StringComparison.Ordinal)));

        Assert.NotNull(resultsGrid);
    }

    [Fact]
    public void PermissionsView_LoadingOverlay_IsInsideRoleGroupResultsGrid()
    {
        var document = LoadViewDocument("PermissionsView.xaml");

        var resultsGrid = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Grid" &&
                string.Equals((string?)candidate.Attribute("Grid.Row"), "3", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Grid.RowSpan"), "2", StringComparison.Ordinal) &&
                candidate.Elements().Any(child =>
                    child.Name.LocalName == "Border" &&
                    string.Equals((string?)child.Attribute("Visibility"), "{Binding Permissions.IsLoading, Converter={StaticResource BoolToVisibility}}", StringComparison.Ordinal)));

        Assert.NotNull(resultsGrid);
    }

    private static XDocument LoadViewDocument(string viewFileName)
    {
        var viewPath = TestPathHelper.GetRepositoryPath("src", "ExchangeAdmin.Presentation", "Views", viewFileName);

        return XDocument.Load(viewPath);
    }
}
