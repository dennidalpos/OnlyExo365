using OnlyExo365.Shell.Security;
using OnlyExo365.Contracts;
using OnlyExo365.Shell.Localization;
using OnlyExo365.Shell.Services;
using OnlyExo365.Shell.Text;
using OnlyExo365.Shell.ViewModels;
using System.Text.RegularExpressions;

namespace OnlyExo365.Tests;

public sealed class UiTextCatalogLocalizationTests : IDisposable
{
    private readonly string _originalLocale;

    public UiTextCatalogLocalizationTests()
    {
        _originalLocale = LocalizationService.Instance.CurrentLocale;
        LocalizationService.Instance.SetLocale("en");
    }

    public void Dispose()
    {
        LocalizationService.Instance.SetLocale(_originalLocale);
    }

    [Theory]
    [InlineData(NavigationPage.Dashboard)]
    [InlineData(NavigationPage.Contacts)]
    [InlineData(NavigationPage.Resources)]
    [InlineData(NavigationPage.PublicFolders)]
    [InlineData(NavigationPage.MobileDevices)]
    [InlineData(NavigationPage.Migration)]
    [InlineData(NavigationPage.Permissions)]
    [InlineData(NavigationPage.Mailboxes)]
    [InlineData(NavigationPage.DeletedMailboxes)]
    [InlineData(NavigationPage.MailboxSpace)]
    [InlineData(NavigationPage.MailboxAccessReport)]
    [InlineData(NavigationPage.DistributionLists)]
    [InlineData(NavigationPage.MessageTrace)]
    [InlineData(NavigationPage.Compliance)]
    [InlineData(NavigationPage.MailSecurity)]
    [InlineData(NavigationPage.MailFlow)]
    [InlineData(NavigationPage.Tools)]
    [InlineData(NavigationPage.Logs)]
    public void GetNavigationLabel_English_ReturnsNonEmptyString(NavigationPage page)
    {
        LocalizationService.Instance.SetLocale("en");

        var label = UiTextCatalog.GetNavigationLabel(page);

        Assert.False(string.IsNullOrWhiteSpace(label));
    }

    [Theory]
    [InlineData(NavigationPage.Dashboard)]
    [InlineData(NavigationPage.Contacts)]
    [InlineData(NavigationPage.Resources)]
    [InlineData(NavigationPage.PublicFolders)]
    [InlineData(NavigationPage.MobileDevices)]
    [InlineData(NavigationPage.Migration)]
    [InlineData(NavigationPage.Permissions)]
    [InlineData(NavigationPage.Mailboxes)]
    [InlineData(NavigationPage.DeletedMailboxes)]
    [InlineData(NavigationPage.MailboxSpace)]
    [InlineData(NavigationPage.MailboxAccessReport)]
    [InlineData(NavigationPage.DistributionLists)]
    [InlineData(NavigationPage.MessageTrace)]
    [InlineData(NavigationPage.Compliance)]
    [InlineData(NavigationPage.MailSecurity)]
    [InlineData(NavigationPage.MailFlow)]
    [InlineData(NavigationPage.Tools)]
    [InlineData(NavigationPage.Logs)]
    public void GetNavigationLabel_Italian_ReturnsNonEmptyString(NavigationPage page)
    {
        LocalizationService.Instance.SetLocale("it");

        var label = UiTextCatalog.GetNavigationLabel(page);

        Assert.False(string.IsNullOrWhiteSpace(label));
    }

    [Theory]
    [InlineData(NavigationPage.Dashboard)]
    [InlineData(NavigationPage.Contacts)]
    [InlineData(NavigationPage.Mailboxes)]
    [InlineData(NavigationPage.DistributionLists)]
    [InlineData(NavigationPage.Compliance)]
    [InlineData(NavigationPage.Tools)]
    [InlineData(NavigationPage.Logs)]
    public void GetNavigationLabel_NeverReturnsRawKey(NavigationPage page)
    {
        foreach (var locale in new[] { "en", "it" })
        {
            LocalizationService.Instance.SetLocale(locale);
            var label = UiTextCatalog.GetNavigationLabel(page);

            Assert.DoesNotContain("Nav.", label);
        }
    }

    [Fact]
    public void GetNavigationLabel_Italian_DiffersFromEnglish_ForTranslatedPages()
    {
        var translatedPages = new[]
        {
            NavigationPage.Contacts,
            NavigationPage.Mailboxes,
            NavigationPage.DeletedMailboxes,
            NavigationPage.DistributionLists,
            NavigationPage.PublicFolders,
            NavigationPage.Resources,
            NavigationPage.Compliance,
            NavigationPage.Tools,
            NavigationPage.Logs
        };

        foreach (var page in translatedPages)
        {
            LocalizationService.Instance.SetLocale("en");
            var english = UiTextCatalog.GetNavigationLabel(page);

            LocalizationService.Instance.SetLocale("it");
            var italian = UiTextCatalog.GetNavigationLabel(page);

            Assert.True(english != italian,
                $"Expected Italian label for {page} to differ from English label '{english}'");
        }
    }

    [Fact]
    public void AllNavigationPages_CoveredByGetNavigationLabel()
    {
        LocalizationService.Instance.SetLocale("en");

        foreach (NavigationPage page in Enum.GetValues<NavigationPage>())
        {
            var label = UiTextCatalog.GetNavigationLabel(page);
            Assert.False(string.IsNullOrWhiteSpace(label),
                $"GetNavigationLabel returned empty for {page}");
        }
    }

    [Theory]
    [InlineData("Compliance.QueueSearchBtn")]
    [InlineData("Compliance.AuditDateRangeError")]
    [InlineData("Compliance.WorkspaceLoadError")]
    [InlineData("Compliance.SearchGroup")]
    [InlineData("Compliance.CasesActionsGroup")]
    [InlineData("Compliance.RefreshWorkspaceBtn")]
    [InlineData("Compliance.CreateSearchBtn")]
    [InlineData("Compliance.StartSearchBtn")]
    [InlineData("Compliance.RemoveSearchBtn")]
    [InlineData("Compliance.RunPurgeBtn")]
    [InlineData("Compliance.CreateHoldBtn")]
    [InlineData("Compliance.ColName")]
    [InlineData("Compliance.ColCase")]
    [InlineData("Compliance.ColStatus")]
    [InlineData("Compliance.ColLocations")]
    [InlineData("Compliance.ColQuery")]
    [InlineData("Compliance.ColType")]
    [InlineData("Compliance.ColActionType")]
    [InlineData("Compliance.ColSearch")]
    [InlineData("Compliance.ColCreated")]
    [InlineData("Compliance.ColDetails")]
    [InlineData("Compliance.AuditColDate")]
    [InlineData("Compliance.AuditColUser")]
    [InlineData("Compliance.AuditColOperation")]
    [InlineData("Compliance.AuditColObject")]
    [InlineData("Compliance.AuditColRecordType")]
    [InlineData("Compliance.AuditColAuditData")]
    [InlineData("Contact.FieldPassword")]
    [InlineData("Contact.OptionsGroup")]
    [InlineData("Contact.HideFromAddressLists")]
    [InlineData("Distribution.IncludeDynamic")]
    [InlineData("Distribution.NewGroupHeader")]
    [InlineData("Distribution.ColDisplayName")]
    [InlineData("Distribution.ColEmail")]
    [InlineData("Distribution.ColType")]
    [InlineData("Distribution.ColRecipientTypeDetails")]
    [InlineData("Distribution.BackBtn")]
    [InlineData("Distribution.OwnersGroup")]
    [InlineData("Distribution.GroupDetailsGroup")]
    [InlineData("Dashboard.ColName")]
    [InlineData("Dashboard.ColUpn")]
    [InlineData("Dashboard.ColRole")]
    [InlineData("MailboxSettings.PrimarySmtp")]
    [InlineData("MailboxSettings.PrimarySmtpTooltip")]
    [InlineData("MailboxSettings.ProxyAddresses")]
    [InlineData("MailboxSettings.ProxyAddressesTooltip")]
    [InlineData("MailboxSettings.ProxyAddressesHint")]
    [InlineData("MailboxSettings.ForwardingAddress")]
    [InlineData("MailboxSettings.ForwardingAddressTooltip")]
    [InlineData("MailboxSettings.ForwardingSmtp")]
    [InlineData("MailboxSettings.ForwardingSmtpTooltip")]
    [InlineData("MailboxSettings.DeliverToMailboxLabel")]
    [InlineData("MailboxSettings.DeliverToMailboxTooltip")]
    [InlineData("MailboxSettings.MaxSendLabel")]
    [InlineData("MailboxSettings.MaxSendTooltip")]
    [InlineData("MailboxSettings.MaxReceiveLabel")]
    [InlineData("MailboxSettings.MaxReceiveTooltip")]
    [InlineData("MailboxSettings.SizeFormatHint")]
    [InlineData("MailSecurity.RefreshWorkspaceBtn")]
    public void NewLocalizationKeys_ArePresentInEnglishAndItalian(string key)
    {
        foreach (var locale in new[] { "en", "it" })
        {
            LocalizationService.Instance.SetLocale(locale);

            var value = Loc.Get(key);

            Assert.False(string.IsNullOrWhiteSpace(value));
            Assert.NotEqual(key, value);
        }
    }

    [Fact]
    public void XamlLocalizationKeys_ArePresentInEnglishAndItalian()
    {
        var viewsPath = TestPathHelper.GetRepositoryPath("src", "OnlyExo365.Shell", "Views");
        var keyPattern = new Regex(@"loc:Loc\s+Key=(?<key>[A-Za-z0-9_.]+)", RegexOptions.Compiled);
        var keys = Directory.EnumerateFiles(viewsPath, "*.xaml")
            .SelectMany(path => keyPattern.Matches(File.ReadAllText(path)).Select(match => match.Groups["key"].Value))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(key => key, StringComparer.Ordinal)
            .ToArray();

        Assert.NotEmpty(keys);

        foreach (var key in keys)
        {
            foreach (var locale in new[] { "en", "it" })
            {
                LocalizationService.Instance.SetLocale(locale);

                var value = Loc.Get(key);

                Assert.False(string.IsNullOrWhiteSpace(value), $"Missing value for '{key}' in '{locale}'.");
                Assert.NotEqual(key, value);
            }
        }
    }

    [Fact]
    public void ItalianResource_DoesNotContainMojibakeMarkers()
    {
        var resourcePath = TestPathHelper.GetRepositoryPath(
            "src",
            "OnlyExo365.Shell",
            "Localization",
            "Strings.it.resx");

        var content = File.ReadAllText(resourcePath);

        Assert.DoesNotContain("Ã", content, StringComparison.Ordinal);
        Assert.DoesNotContain("Â", content, StringComparison.Ordinal);
        Assert.Contains("Conformità", content, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Status.Worker.Running", "In esecuzione")]
    [InlineData("Status.Connection.Disconnected", "Disconnesso")]
    [InlineData("Tools.Capabilities.NotDetected", "Non rilevate")]
    [InlineData("Common.Yes", "Sì")]
    [InlineData("Tools.Catalog.UpToDateWithEntries", "Aggiornato (2026.04, 1250 voci)", "2026.04", 1250)]
    [InlineData("Tools.AutoUpdate.Daily", "Giornaliero")]
    public void ToolsDynamicKeys_Italian_ReturnsLocalizedText(string key, string expected, params object[] args)
    {
        LocalizationService.Instance.SetLocale("it");

        var value = args.Length == 0
            ? Loc.Get(key)
            : Loc.GetFormat(key, args);

        Assert.Equal(expected, value);
    }

    [Fact]
    public void ToolsView_NoLongerEmbedsObservedEnglishDynamicStates()
    {
        var viewPath = TestPathHelper.GetRepositoryPath(
            "src",
            "OnlyExo365.Shell",
            "Views",
            "ToolsView.xaml");

        var content = File.ReadAllText(viewPath);

        Assert.DoesNotContain("Value=\"Disconnected\"", content, StringComparison.Ordinal);
        Assert.DoesNotContain("Value=\"Connected\"", content, StringComparison.Ordinal);
        Assert.DoesNotContain("Value=\"Not detected\"", content, StringComparison.Ordinal);
        Assert.DoesNotContain("Value=\"Yes\"", content, StringComparison.Ordinal);
        Assert.DoesNotContain("Value=\"No\"", content, StringComparison.Ordinal);
        Assert.Contains("Text=\"{Binding ExchangeStateDisplay}\"", content, StringComparison.Ordinal);
        Assert.Contains("Text=\"{Binding CapabilitiesDisplay}\"", content, StringComparison.Ordinal);
        Assert.Contains("Text=\"{Binding WorkerRunningDisplay}\"", content, StringComparison.Ordinal);
    }

    [Fact]
    public void LeastPrivilegeDisplay_Italian_LocalizesVisiblePendingRows()
    {
        LocalizationService.Instance.SetLocale("it");
        var evaluator = new LeastPrivilegeEvaluator(ExchangeOnlineConfiguration.CreateDefault());

        var firstRow = new LeastPrivilegeFeatureDisplay(evaluator.EvaluateAll(null)[0]);

        Assert.Equal("Inventario e azioni ActiveSync", firstRow.FeatureName);
        Assert.Equal("In attesa", firstRow.StatusLabel);
        Assert.Contains("Connettiti a Exchange", firstRow.ValidationMessage, StringComparison.Ordinal);
        Assert.DoesNotContain("Inventory device", firstRow.Description, StringComparison.Ordinal);
    }
}

