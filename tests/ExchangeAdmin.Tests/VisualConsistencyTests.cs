using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ExchangeAdmin.Tests;

public sealed class VisualConsistencyTests
{
    [Fact]
    public void DarkTheme_DefinesSharedVisualConsistencyStyles()
    {
        var document = LoadTheme("DarkTheme.xaml");
        var styleKeys = document
            .Descendants()
            .Where(candidate => candidate.Name.LocalName == "Style")
            .Select(candidate => (string?)candidate.Attribute(XName.Get("Key", "http://schemas.microsoft.com/winfx/2006/xaml")))
            .Where(static key => !string.IsNullOrWhiteSpace(key))
            .ToHashSet(StringComparer.Ordinal);

        Assert.Contains("PageTitleTextStyle", styleKeys);
        Assert.Contains("SectionHeadingTextStyle", styleKeys);
        Assert.Contains("CaptionTextStyle", styleKeys);
        Assert.Contains("ErrorBannerBorderStyle", styleKeys);
        Assert.Contains("WarningBannerBorderStyle", styleKeys);
        Assert.Contains("InfoBannerBorderStyle", styleKeys);
        Assert.Contains("InlineValidationTextStyle", styleKeys);
        Assert.Contains("ActionGroupPrimaryButton", styleKeys);
        Assert.Contains("CompactErrorOverlayCardStyle", styleKeys);
        Assert.Contains("StatusChipBorderStyle", styleKeys);
    }

    [Fact]
    public void DarkTheme_UsesDarkComboBoxPalette()
    {
        var document = LoadTheme("DarkTheme.xaml");

        var comboBoxStyle = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Style" &&
                string.Equals((string?)candidate.Attribute("TargetType"), "ComboBox", StringComparison.Ordinal));

        Assert.NotNull(comboBoxStyle);
        Assert.Contains(comboBoxStyle!.Elements(), element =>
            element.Name.LocalName == "Setter" &&
            string.Equals((string?)element.Attribute("Property"), "Background", StringComparison.Ordinal) &&
            string.Equals((string?)element.Attribute("Value"), "{StaticResource TertiaryBackgroundBrush}", StringComparison.Ordinal));
        Assert.Contains(comboBoxStyle.Elements(), element =>
            element.Name.LocalName == "Setter" &&
            string.Equals((string?)element.Attribute("Property"), "Foreground", StringComparison.Ordinal) &&
            string.Equals((string?)element.Attribute("Value"), "{StaticResource PrimaryTextBrush}", StringComparison.Ordinal));
    }

    [Fact]
    public void DarkTheme_GroupBoxTemplate_UsesThemeBorderAndHeaderTreatment()
    {
        var document = LoadTheme("DarkTheme.xaml");
        var groupBoxStyle = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Style" &&
                string.Equals((string?)candidate.Attribute("TargetType"), "GroupBox", StringComparison.Ordinal));

        Assert.NotNull(groupBoxStyle);
        Assert.Contains(groupBoxStyle!.Descendants(), candidate =>
            candidate.Name.LocalName == "ContentPresenter" &&
            string.Equals((string?)candidate.Attribute("ContentSource"), "Header", StringComparison.Ordinal));
        Assert.Contains(groupBoxStyle.Descendants(), candidate =>
            candidate.Name.LocalName == "Border" &&
            string.Equals((string?)candidate.Attribute("BorderBrush"), "{TemplateBinding BorderBrush}", StringComparison.Ordinal) &&
            string.Equals((string?)candidate.Attribute("CornerRadius"), "4", StringComparison.Ordinal));
    }

    [Fact]
    public void DashboardView_UsesSharedBannerAndOverlayStyles()
    {
        var document = LoadView("DashboardView.xaml");

        Assert.Contains(document.Descendants(), candidate =>
            candidate.Name.LocalName == "TextBlock" &&
            string.Equals((string?)candidate.Attribute("Text"), "{loc:Loc Key=Dashboard.Title}", StringComparison.Ordinal) &&
            string.Equals((string?)candidate.Attribute("Style"), "{StaticResource PageTitleTextStyle}", StringComparison.Ordinal));

        Assert.Contains(document.Descendants(), candidate =>
            candidate.Name.LocalName == "Border" &&
            string.Equals((string?)candidate.Attribute("Style"), "{StaticResource WarningBannerBorderStyle}", StringComparison.Ordinal));

        Assert.Contains(document.Descendants(), candidate =>
            candidate.Name.LocalName == "Border" &&
            string.Equals((string?)candidate.Attribute("Style"), "{StaticResource LoadingOverlayBackdropBorderStyle}", StringComparison.Ordinal));
    }

    [Fact]
    public void DashboardView_UsesLocalizedLicenseCounterLabels()
    {
        var document = LoadView("DashboardView.xaml");
        var labelTexts = document
            .Descendants()
            .Where(candidate => candidate.Name.LocalName == "TextBlock")
            .Select(candidate => (string?)candidate.Attribute("Text"))
            .Where(static text => !string.IsNullOrWhiteSpace(text))
            .ToHashSet(StringComparer.Ordinal);

        Assert.Contains("{loc:Loc Key=Dashboard.LicensesTotal}", labelTexts);
        Assert.Contains("{loc:Loc Key=Dashboard.LicensesAssigned}", labelTexts);
        Assert.Contains("{loc:Loc Key=Dashboard.LicensesAvailable}", labelTexts);
    }

    [Fact]
    public void MailFlowView_UsesSharedActionGroupAndValidationStyles()
    {
        var document = LoadView("MailFlowView.xaml");

        Assert.Contains(document.Descendants(), candidate =>
            candidate.Name.LocalName == "Button" &&
            string.Equals((string?)candidate.Attribute("Content"), "{loc:Loc Key=MailFlow.NewReset}", StringComparison.Ordinal) &&
            string.Equals((string?)candidate.Attribute("Style"), "{StaticResource ActionGroupPrimaryButton}", StringComparison.Ordinal));

        Assert.Contains(document.Descendants(), candidate =>
            candidate.Name.LocalName == "TextBlock" &&
            string.Equals((string?)candidate.Attribute("Style"), "{StaticResource InlineValidationTextStyle}", StringComparison.Ordinal));

        Assert.Contains(document.Descendants(), candidate =>
            candidate.Name.LocalName == "Border" &&
            string.Equals((string?)candidate.Attribute("Style"), "{StaticResource LoadingOverlayBackdropBorderStyle}", StringComparison.Ordinal) &&
            string.Equals((string?)candidate.Attribute("Visibility"), "{Binding MailFlow.IsLoading, Converter={StaticResource BoolToVisibility}}", StringComparison.Ordinal));
    }

    [Fact]
    public void ComplianceView_UsesSharedSelectedSearchActionStyles()
    {
        var document = LoadView("ComplianceView.xaml");

        AssertButtonStyle(document, "{loc:Loc Key=Compliance.StartSearchBtn}", "{StaticResource PrimaryButton}");
        AssertButtonStyle(document, "{loc:Loc Key=Compliance.RemoveSearchBtn}", "{StaticResource DangerButton}");
        AssertButtonStyle(document, "{loc:Loc Key=Compliance.RunPurgeBtn}", "{StaticResource PrimaryButton}");
        AssertButtonStyle(document, "{loc:Loc Key=Compliance.CreateHoldBtn}", "{StaticResource PrimaryButton}");
    }

    [Fact]
    public void DistributionListView_UsesSharedActionStylesAndLocalizedLabels()
    {
        var document = LoadView("DistributionListView.xaml");
        var content = File.ReadAllText(TestPathHelper.GetRepositoryPath("src", "ExchangeAdmin.Presentation", "Views", "DistributionListView.xaml"));

        Assert.Contains("Content=\"{loc:Loc Key=Distribution.BackBtn}\"", content, StringComparison.Ordinal);
        Assert.Contains("Content=\"{loc:Loc Key=Distribution.IncludeDynamic}\"", content, StringComparison.Ordinal);
        Assert.DoesNotContain("Indietro", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Includi", content, StringComparison.OrdinalIgnoreCase);

        AssertButtonStyle(document, "{loc:Loc Key=Distribution.BackBtn}", "{StaticResource InlineActionButton}");
        AssertButtonStyle(document, "{loc:Loc Key=Btn.Cancel}", "{StaticResource InlineActionButton}");
        AssertButtonStyle(document, "{loc:Loc Key=Distribution.MemberPreview}", "{StaticResource InlineActionButton}");
        AssertButtonStyle(document, "{loc:Loc Key=Distribution.AddMember}", "{StaticResource InlineActionButton}");
        AssertButtonStyle(document, "{loc:Loc Key=Btn.LoadMore}", "{StaticResource InlineActionButton}");
        AssertButtonStyle(document, "{loc:Loc Key=Distribution.LoadMoreMembers}", "{StaticResource InlineActionButton}");
        AssertButtonCommandStyle(document, "{Binding DistributionLists.AddAcceptSenderCommand}", "{StaticResource InlineActionButton}");
        AssertButtonCommandStyle(document, "{Binding DistributionLists.AddRejectSenderCommand}", "{StaticResource InlineActionButton}");
    }

    [Fact]
    public void MailboxDetailsView_UsesSharedSavingOverlayBackdrop()
    {
        var document = LoadView("MailboxDetailsView.xaml");

        Assert.Contains(document.Descendants(), candidate =>
            candidate.Name.LocalName == "Border" &&
            string.Equals((string?)candidate.Attribute("Style"), "{StaticResource LoadingOverlayBackdropBorderStyle}", StringComparison.Ordinal) &&
            string.Equals((string?)candidate.Attribute("Visibility"), "{Binding IsSaving, Converter={StaticResource BoolToVisibility}}", StringComparison.Ordinal));
        Assert.DoesNotContain(document.Descendants(), candidate =>
            candidate.Name.LocalName == "Border" &&
            string.Equals((string?)candidate.Attribute("Background"), "#CC000000", StringComparison.Ordinal));
        AssertButtonCommandStyle(document, "{Binding DiscardMailboxChangesCommand}", "{StaticResource InlineActionButton}");
        AssertButtonCommandStyle(document, "{Binding DiscardPermissionsCommand}", "{StaticResource InlineActionButton}");
    }

    [Fact]
    public void MailboxPermissionsTab_UsesSharedCaptionStyleForFieldLabels()
    {
        var document = LoadView("MailboxPermissionsTab.xaml");

        AssertTextStyle(document, "{loc:Loc Key=MailboxPerms.FieldUser}", "{StaticResource CaptionTextStyle}");
        AssertTextStyle(document, "{loc:Loc Key=MailboxPerms.FieldPermissionType}", "{StaticResource CaptionTextStyle}");
        AssertTextStyle(document, "{loc:Loc Key=MailboxPerms.FieldAutoMap}", "{StaticResource CaptionTextStyle}");
        AssertTextStyle(document, "{loc:Loc Key=MailboxPerms.FieldFolderPath}", "{StaticResource CaptionTextStyle}");
        AssertTextStyle(document, "{loc:Loc Key=MailboxPerms.FieldResolvedTarget}", "{StaticResource CaptionTextStyle}");
    }

    [Fact]
    public void PermissionsView_UsesSharedActionStylesAndThemeBackgrounds()
    {
        var document = LoadView("PermissionsView.xaml");
        var content = File.ReadAllText(TestPathHelper.GetRepositoryPath("src", "ExchangeAdmin.Presentation", "Views", "PermissionsView.xaml"));

        Assert.DoesNotContain("Copy da", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("iniziali", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("#33000000", content, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("{loc:Loc Key=Permissions.LoadingDetails}", content, StringComparison.Ordinal);
        AssertButtonCommandStyle(document, "{Binding Permissions.NewRoleGroupCommand}", "{StaticResource InlineActionButton}");
        AssertButtonStyle(document, "{loc:Loc Key=Permissions.Reset}", "{StaticResource InlineActionButton}");
    }

    [Fact]
    public void MailboxReportingViews_UseLocalizedLabelsAndSharedExportStyles()
    {
        var accessReport = LoadView("MailboxAccessReportView.xaml");
        var mailboxSpace = LoadView("MailboxSpaceView.xaml");

        AssertButtonCommandStyle(accessReport, "{Binding MailboxAccessReport.ExportExcelCommand}", "{StaticResource InlineActionButton}");
        AssertButtonStyle(accessReport, "{loc:Loc Key=MailboxAccess.StartAnalysis}", "{StaticResource PrimaryButton}");
        AssertButtonCommandStyle(mailboxSpace, "{Binding MailboxSpace.ExportExcelCommand}", "{StaticResource InlineActionButton}");
        AssertButtonStyle(mailboxSpace, "{loc:Loc Key=MailboxSpace.StartScan}", "{StaticResource PrimaryButton}");
    }

    [Theory]
    [InlineData("DeletedMailboxesView.xaml", "Load list")]
    [InlineData("DeletedMailboxesView.xaml", "Check UPN")]
    [InlineData("MailboxAccessReportView.xaml", "Mailbox access summary")]
    [InlineData("MailboxAccessReportView.xaml", "Start analysis")]
    [InlineData("MailboxSpaceView.xaml", "Sorted by % free space")]
    [InlineData("MailboxSpaceView.xaml", "Start scan")]
    [InlineData("MessageTraceView.xaml", "Export Excel")]
    [InlineData("MessageTraceView.xaml", "Message details")]
    [InlineData("MailboxOverviewTab.xaml", "Basic information")]
    [InlineData("MailboxRestoreTab.xaml", "Source mailbox (UPN/GUID):")]
    [InlineData("MailboxLicensesTab.xaml", "Select a license to assign")]
    [InlineData("DistributionListView.xaml", "Email addresses")]
    [InlineData("DistributionListView.xaml", "Allow external senders")]
    [InlineData("LogsView.xaml", "Filter logs by minimum level")]
    [InlineData("MigrationView.xaml", "Created: {0:g}")]
    [InlineData("MigrationView.xaml", "AutoStart: {0}")]
    [InlineData("ToolsView.xaml", "Module: {0}")]
    [InlineData("ToolsView.xaml", "Graph scopes: {0}")]
    public void TargetedUxViews_DoNotUseHardcodedOperatorCopy(string fileName, string hardcodedCopy)
    {
        var content = File.ReadAllText(TestPathHelper.GetRepositoryPath("src", "ExchangeAdmin.Presentation", "Views", fileName));

        Assert.DoesNotContain(hardcodedCopy, content, StringComparison.Ordinal);
    }

    [Fact]
    public void Views_DoNotUseHardcodedStringFormatPrefixes()
    {
        var viewsPath = TestPathHelper.GetRepositoryPath("src", "ExchangeAdmin.Presentation", "Views");
        var offendingLines = Directory.EnumerateFiles(viewsPath, "*.xaml")
            .SelectMany(path => File.ReadLines(path)
                .Select((line, index) => new { Path = path, LineNumber = index + 1, Line = line }))
            .Where(candidate => Regex.IsMatch(candidate.Line, @"StringFormat=[A-Za-z][^}""]*:"))
            .Select(candidate => $"{Path.GetFileName(candidate.Path)}:{candidate.LineNumber}: {candidate.Line.Trim()}")
            .ToArray();

        Assert.Empty(offendingLines);
    }

    [Fact]
    public void MobileDevicesView_UsesSharedActionStyles()
    {
        var document = LoadView("MobileDevicesView.xaml");

        AssertButtonStyle(document, "{loc:Loc Key=MobileDevices.AssignPolicy}", "{StaticResource PrimaryButton}");
        AssertButtonStyle(document, "{loc:Loc Key=MobileDevices.Block}", "{StaticResource DangerButton}");
        AssertButtonStyle(document, "{loc:Loc Key=MobileDevices.Quarantine}", "{StaticResource InlineActionButton}");
    }

    [Fact]
    public void DeletedMailboxesView_ReliesOnSharedGridAndOverlayStyles()
    {
        var document = LoadView("DeletedMailboxesView.xaml");

        Assert.DoesNotContain(document.Descendants(), candidate => candidate.Name.LocalName == "DataGrid.RowStyle");
        Assert.DoesNotContain(document.Descendants(), candidate => candidate.Name.LocalName == "DataGrid.ColumnHeaderStyle");
        Assert.Contains(document.Descendants(), candidate =>
            candidate.Name.LocalName == "Border" &&
            string.Equals((string?)candidate.Attribute("Style"), "{StaticResource CompactErrorOverlayCardStyle}", StringComparison.Ordinal));
    }

    [Fact]
    public void LogsView_UsesSharedCaptionAndStatusChipStyles()
    {
        var document = LoadView("LogsView.xaml");

        Assert.Contains(document.Descendants(), candidate =>
            candidate.Name.LocalName == "TextBlock" &&
            string.Equals((string?)candidate.Attribute("Text"), "{Binding Logs.PersistentLogPolicySummary}", StringComparison.Ordinal) &&
            string.Equals((string?)candidate.Attribute("Style"), "{StaticResource CaptionTextStyle}", StringComparison.Ordinal));

        Assert.Contains(document.Descendants(), candidate =>
            candidate.Name.LocalName == "Border" &&
            string.Equals((string?)candidate.Attribute("Style"), "{StaticResource StatusChipBorderStyle}", StringComparison.Ordinal));
    }

    [Fact]
    public void LogsView_DoesNotOverrideSharedActionAndChipDensity()
    {
        var document = LoadView("LogsView.xaml");

        AssertButtonStyle(document, "{loc:Loc Key=Btn.Clear}", "{StaticResource ActionGroupDangerButton}");
        Assert.Contains(document.Descendants(), candidate =>
            candidate.Name.LocalName == "StackPanel" &&
            string.Equals((string?)candidate.Attribute("Grid.Column"), "1", StringComparison.Ordinal) &&
            string.Equals((string?)candidate.Attribute("Orientation"), "Horizontal", StringComparison.Ordinal) &&
            string.Equals((string?)candidate.Attribute("VerticalAlignment"), "Top", StringComparison.Ordinal));
        Assert.DoesNotContain(document.Descendants(), candidate =>
            candidate.Name.LocalName == "Border" &&
            string.Equals((string?)candidate.Attribute("Style"), "{StaticResource StatusChipBorderStyle}", StringComparison.Ordinal) &&
            candidate.Attribute("CornerRadius") is not null);
        Assert.DoesNotContain(document.Descendants(), candidate =>
            candidate.Name.LocalName == "Setter" &&
            string.Equals((string?)candidate.Attribute("Property"), "Background", StringComparison.Ordinal) &&
            ((string?)candidate.Attribute("Value") is "#2D0000" or "#2D2600"));
    }

    [Fact]
    public void ToolsView_DoesNotOverrideSharedBannerStatusAndActionDensity()
    {
        var document = LoadView("ToolsView.xaml");

        Assert.DoesNotContain(document.Descendants(), candidate =>
            candidate.Name.LocalName == "Border" &&
            ((string?)candidate.Attribute("Style") is "{StaticResource ErrorBannerBorderStyle}" or "{StaticResource WarningBannerBorderStyle}") &&
            (candidate.Attribute("Padding") is not null || candidate.Attribute("Margin") is not null));
        Assert.DoesNotContain(document.Descendants(), candidate =>
            candidate.Name.LocalName == "TextBlock" &&
            string.Equals((string?)candidate.Attribute("Style"), "{StaticResource ErrorBannerTextStyle}", StringComparison.Ordinal) &&
            candidate.Attribute("FontSize") is not null);
        Assert.DoesNotContain(document.Descendants(), candidate =>
            candidate.Name.LocalName == "Ellipse" &&
            string.Equals((string?)candidate.Attribute("Width"), "10", StringComparison.Ordinal) &&
            string.Equals((string?)candidate.Attribute("Height"), "10", StringComparison.Ordinal));
        Assert.DoesNotContain(document.Descendants(), candidate =>
            candidate.Name.LocalName == "Button" &&
            string.Equals((string?)candidate.Attribute("Style"), "{StaticResource ActionGroupButton}", StringComparison.Ordinal) &&
            candidate.Attribute("Padding") is not null);
        Assert.DoesNotContain(document.Descendants(), candidate =>
            candidate.Name.LocalName == "Button" &&
            string.Equals((string?)candidate.Attribute("Style"), "{StaticResource ActionGroupButton}", StringComparison.Ordinal) &&
            candidate.Attribute("Margin") is not null);
        Assert.DoesNotContain(document.Descendants(), candidate =>
            candidate.Name.LocalName == "Border" &&
            string.Equals((string?)candidate.Attribute("Style"), "{StaticResource StatusChipBorderStyle}", StringComparison.Ordinal) &&
            candidate.Attribute("Padding") is not null);
    }

    [Theory]
    [InlineData("ComplianceView.xaml")]
    [InlineData("DashboardView.xaml")]
    [InlineData("DistributionListView.xaml")]
    [InlineData("MailboxAccessReportView.xaml")]
    [InlineData("MailboxRestoreTab.xaml")]
    [InlineData("MailboxSettingsTab.xaml")]
    [InlineData("MailboxSpaceView.xaml")]
    [InlineData("MailboxPermissionsTab.xaml")]
    [InlineData("MobileDevicesView.xaml")]
    [InlineData("PermissionsView.xaml")]
    public void Views_DoNotContainKnownMixedLanguageLabels(string fileName)
    {
        var content = File.ReadAllText(TestPathHelper.GetRepositoryPath("src", "ExchangeAdmin.Presentation", "Views", fileName));

        Assert.DoesNotContain("Totale", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Assegnate", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Disponibili", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Accesso", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Invia", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Ereditato", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("autorizz", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Indietro", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Includi", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Gli hold", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Avvia", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Scansione", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Usato", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Residuo", content, StringComparison.OrdinalIgnoreCase);
        AssertDoesNotContainItalianWord(content, "Limite");
        Assert.DoesNotContain("Archivio", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Recupero", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Assegna", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Consenti", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Azione", content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Avanzamento", content, StringComparison.OrdinalIgnoreCase);
    }

    private static void AssertButtonStyle(XDocument document, string content, string expectedStyle)
    {
        var button = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Content"), content, StringComparison.Ordinal));

        Assert.NotNull(button);
        Assert.Equal(expectedStyle, (string?)button!.Attribute("Style"));
    }

    private static void AssertButtonCommandStyle(XDocument document, string command, string expectedStyle)
    {
        var button = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Command"), command, StringComparison.Ordinal));

        Assert.NotNull(button);
        Assert.Equal(expectedStyle, (string?)button!.Attribute("Style"));
    }

    private static void AssertTextStyle(XDocument document, string text, string expectedStyle)
    {
        var textBlock = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "TextBlock" &&
                string.Equals((string?)candidate.Attribute("Text"), text, StringComparison.Ordinal));

        Assert.NotNull(textBlock);
        Assert.Equal(expectedStyle, (string?)textBlock!.Attribute("Style"));
    }

    private static void AssertDoesNotContainItalianWord(string content, string word)
    {
        var pattern = $@"(?<![A-Za-z]){Regex.Escape(word)}(?![A-Za-z])";
        Assert.False(Regex.IsMatch(content, pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant));
    }

    private static XDocument LoadTheme(string fileName)
        => XDocument.Load(TestPathHelper.GetRepositoryPath("src", "ExchangeAdmin.Presentation", "Themes", fileName));

    private static XDocument LoadView(string fileName)
        => XDocument.Load(TestPathHelper.GetRepositoryPath("src", "ExchangeAdmin.Presentation", "Views", fileName));
}
