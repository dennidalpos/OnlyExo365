using System.Xml.Linq;

namespace OnlyExo365.Tests;

public sealed class DataGridLayoutTests
{
    [Fact]
    public void DarkTheme_DataGridStyle_EnablesAutosizeAndManualResize()
    {
        var document = LoadThemeDocument("DarkTheme.xaml");

        var style = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Style" &&
                string.Equals((string?)candidate.Attribute("TargetType"), "DataGrid", StringComparison.Ordinal));

        Assert.NotNull(style);
        AssertSetter(style!, "ColumnWidth", "Auto");
        AssertSetter(style!, "MinColumnWidth", "90");
        AssertSetter(style!, "CanUserResizeColumns", "True");
        AssertSetter(style!, "ScrollViewer.HorizontalScrollBarVisibility", "Auto");
    }

    [Fact]
    public void DarkTheme_ScrollBarStyle_UsesOrientationSpecificThickness()
    {
        var document = LoadThemeDocument("DarkTheme.xaml");

        var style = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Style" &&
                string.Equals((string?)candidate.Attribute("TargetType"), "ScrollBar", StringComparison.Ordinal));

        Assert.NotNull(style);

        var triggers = style!
            .Descendants()
            .Where(candidate => candidate.Name.LocalName == "Trigger")
            .ToList();

        Assert.Contains(triggers, trigger =>
            string.Equals((string?)trigger.Attribute("Property"), "Orientation", StringComparison.Ordinal) &&
            string.Equals((string?)trigger.Attribute("Value"), "Vertical", StringComparison.Ordinal) &&
            trigger.Elements().Any(setter =>
                setter.Name.LocalName == "Setter" &&
                string.Equals((string?)setter.Attribute("Property"), "Width", StringComparison.Ordinal) &&
                string.Equals((string?)setter.Attribute("Value"), "6", StringComparison.Ordinal)));

        Assert.Contains(triggers, trigger =>
            string.Equals((string?)trigger.Attribute("Property"), "Orientation", StringComparison.Ordinal) &&
            string.Equals((string?)trigger.Attribute("Value"), "Horizontal", StringComparison.Ordinal) &&
            trigger.Elements().Any(setter =>
                setter.Name.LocalName == "Setter" &&
                string.Equals((string?)setter.Attribute("Property"), "Height", StringComparison.Ordinal) &&
                string.Equals((string?)setter.Attribute("Value"), "6", StringComparison.Ordinal)));
    }

    [Fact]
    public void MailboxListView_PrimaryColumns_UseAutoSizingWithMinimumWidth()
    {
        var document = LoadViewDocument("MailboxListView.xaml");

        AssertColumn(document, "{loc:Loc Key=Mailbox.List.ColDisplayName}", "Auto", "180");
        AssertColumn(document, "{loc:Loc Key=Mailbox.List.ColUPN}", "Auto", "220");
        AssertColumn(document, "{loc:Loc Key=Mailbox.List.ColEmail}", "Auto", "220");
    }

    [Fact]
    public void MessageTraceView_PrimaryColumns_UseAutoSizingWithMinimumWidth()
    {
        var document = LoadViewDocument("MessageTraceView.xaml");

        AssertColumn(document, "{loc:Loc Key=MessageTrace.ColSender}", "Auto", "220");
        AssertColumn(document, "{loc:Loc Key=MessageTrace.ColRecipient}", "Auto", "220");
        AssertColumn(document, "{loc:Loc Key=MessageTrace.ColSubject}", "Auto", "260");
        AssertColumn(document, "{loc:Loc Key=MessageTrace.ColDetail}", "Auto", "220");
        AssertColumn(document, "{loc:Loc Key=MessageTrace.ColData}", "Auto", "220");
    }

    [Fact]
    public void ContactsResourcesAndPublicFolders_UseAutoSizingForOperationalColumns()
    {
        var contacts = LoadViewDocument("ContactsView.xaml");
        AssertColumn(contacts, "{loc:Loc Key=Contact.ColType}", "Auto", "110");
        AssertColumn(contacts, "{loc:Loc Key=Contact.ColHideFromGal}", "Auto", "95");

        var resources = LoadViewDocument("ResourcesView.xaml");
        AssertColumn(resources, "{loc:Loc Key=Resources.Type}", "Auto", "100");
        AssertColumn(resources, "{loc:Loc Key=Resources.HideFromGAL}", "Auto", "95");

        var publicFolders = LoadViewDocument("PublicFoldersView.xaml");
        AssertColumn(publicFolders, "{loc:Loc Key=PublicFolders.Mail}", "Auto", "70");
        AssertColumn(publicFolders, "{loc:Loc Key=PublicFolders.Subfolder}", "Auto", "90");
        AssertColumn(publicFolders, "{loc:Loc Key=PublicFolders.Role}", "Auto", "170");
        AssertColumn(publicFolders, "{loc:Loc Key=PublicFolders.Actions}", "Auto", "160");
    }

    [Fact]
    public void ComplianceMailSecurityAndMailFlow_UseAutoSizingForTenantResultColumns()
    {
        var compliance = LoadViewDocument("ComplianceView.xaml");
        AssertColumn(compliance, "{loc:Loc Key=Compliance.ColTask}", "Auto", "140");
        AssertColumn(compliance, "{loc:Loc Key=Compliance.AuditColAuditData}", "Auto", "260");
        AssertColumn(compliance, "{loc:Loc Key=Compliance.ColQuery}", "Auto", "240");
        AssertColumn(compliance, "{loc:Loc Key=Compliance.ColDetails}", "Auto", "240");

        var mailSecurity = LoadViewDocument("MailSecurityView.xaml");
        AssertColumn(mailSecurity, "{loc:Loc Key=MailSecurity.Domain}", "Auto", "200");
        AssertColumn(mailSecurity, "{loc:Loc Key=MailSecurity.SpamAction}", "Auto", "220");
        AssertColumn(mailSecurity, "{loc:Loc Key=MailSecurity.AuthFail}", "Auto", "220");
        AssertColumn(mailSecurity, "{loc:Loc Key=MailSecurity.Action}", "Auto", "220");

        var mailFlow = LoadViewDocument("MailFlowView.xaml");
        AssertColumn(mailFlow, "{loc:Loc Key=MailFlow.Priority}", "Auto", "90");
        AssertColumn(mailFlow, "{loc:Loc Key=MailFlow.State}", "Auto", "120");
        AssertColumn(mailFlow, "{loc:Loc Key=MailFlow.Recipients}", "Auto", "140");
        AssertColumn(mailFlow, "{loc:Loc Key=MailFlow.DiffRetention}", "Auto", "110");
    }

    private static void AssertColumn(XDocument document, string header, string width, string minWidth)
    {
        var column = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName is "DataGridTextColumn" or "DataGridCheckBoxColumn" or "DataGridTemplateColumn" &&
                string.Equals((string?)candidate.Attribute("Header"), header, StringComparison.Ordinal));

        Assert.NotNull(column);
        Assert.Equal(width, (string?)column!.Attribute("Width"));
        Assert.Equal(minWidth, (string?)column.Attribute("MinWidth"));
    }

    private static void AssertSetter(XElement style, string property, string value)
    {
        var setter = style
            .Elements()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Setter" &&
                string.Equals((string?)candidate.Attribute("Property"), property, StringComparison.Ordinal));

        Assert.NotNull(setter);
        Assert.Equal(value, (string?)setter!.Attribute("Value"));
    }

    private static XDocument LoadViewDocument(string viewFileName)
    {
        var viewPath = TestPathHelper.GetRepositoryPath("src", "OnlyExo365.Shell", "Views", viewFileName);

        return XDocument.Load(viewPath);
    }

    private static XDocument LoadThemeDocument(string themeFileName)
    {
        var themePath = TestPathHelper.GetRepositoryPath("src", "OnlyExo365.Shell", "Themes", themeFileName);

        return XDocument.Load(themePath);
    }
}

