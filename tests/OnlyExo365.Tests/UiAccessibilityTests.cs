using System.Text.RegularExpressions;
using System.Xml.Linq;
using OnlyExo365.Shell.Localization;

namespace OnlyExo365.Tests;

public sealed class UiAccessibilityTests
{
    [Fact]
    public void EditableFields_HaveAccessibleLabelsAndCatalogKeysExist()
    {
        var directory = TestPathHelper.GetRepositoryPath("src", "OnlyExo365.Shell");
        var keys = XDocument.Load(Path.Combine(directory, "Localization", "Strings.resx"))
            .Root!.Elements("data").Select(entry => (string)entry.Attribute("name")!).ToHashSet();
        foreach (var file in Directory.EnumerateFiles(Path.Combine(directory, "Views"), "*.xaml"))
        {
            var source = File.ReadAllText(file);
            foreach (Match match in Regex.Matches(source, @"\{loc:Loc Key=([^}]+)\}"))
                Assert.Contains(match.Groups[1].Value, keys);
            foreach (var field in XDocument.Parse(source).Descendants().Where(element =>
                (element.Name.LocalName is "TextBox" or "ComboBox" or "DatePicker" or "PasswordBox") &&
                (string?)element.Attribute("IsReadOnly") != "True"))
            {
                Assert.True(field.Attribute("AutomationProperties.Name") != null ||
                    field.Attribute("AutomationProperties.LabeledBy") != null, $"{Path.GetFileName(file)}: {field}");
            }
        }
    }

    [Fact]
    public void Catalogs_HaveMatchingKeysAndFormatArguments()
    {
        var directory = TestPathHelper.GetRepositoryPath("src", "OnlyExo365.Shell", "Localization");
        static Dictionary<string, string> Read(string path) => XDocument.Load(path).Root!.Elements("data")
            .ToDictionary(entry => (string)entry.Attribute("name")!, entry => (string)entry.Element("value")!);
        var english = Read(Path.Combine(directory, "Strings.resx"));
        var italian = Read(Path.Combine(directory, "Strings.it.resx"));
        Assert.Equal(english.Keys.Order(), italian.Keys.Order());
        foreach (var (key, text) in english)
        {
            static IEnumerable<string> Arguments(string value) => Regex.Matches(value, @"\{(\d+)(?:[^}]*)\}")
                .Select(match => match.Groups[1].Value).Distinct().Order();
            Assert.Equal(Arguments(text), Arguments(italian[key]));
        }
    }

    [Fact]
    public void FormattedNumbers_UseSelectedLocaleInsteadOfSystemLocale()
    {
        var service = LocalizationService.Instance;
        var original = service.CurrentLocale;
        try
        {
            service.SetLocale("it");
            Assert.Equal("1,5 elementi", service.GetFormat("Mailbox.ItemsFormat", 1.5));
            service.SetLocale("en");
            Assert.Equal("1.5 items", service.GetFormat("Mailbox.ItemsFormat", 1.5));
        }
        finally { service.SetLocale(original); }
    }
}
