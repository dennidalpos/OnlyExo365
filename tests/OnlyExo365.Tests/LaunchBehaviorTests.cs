using System.Xml.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Markup;

namespace OnlyExo365.Tests;

public sealed class LaunchBehaviorTests
{
    [WpfFact]
    public void MainWindow_LanguageBinding_OpensAndReformatsDatesForBothLocales()
    {
        var source = XDocument.Load(TestPathHelper.GetRepositoryPath(
            "src", "OnlyExo365.Shell", "Views", "MainWindow.xaml"));
        var language = source.Root!.Attribute("Language");
        Assert.NotNull(language);
        var window = (Window)XamlReader.Parse(new XElement(source.Root.Name,
            new XAttribute(language!)).ToString());
        window.Left = -10000;
        window.Top = -10000;
        window.WindowStartupLocation = WindowStartupLocation.Manual;
        var date = new DateTime(2026, 10, 10, 15, 30, 0);
        var text = new TextBlock();
        text.SetBinding(TextBlock.TextProperty, new Binding { Source = date, StringFormat = "{0:g}" });
        window.Content = text;
        window.DataContext = new { LanguageSelection = new { SelectedLocale = new { Code = "en" } } };
        try
        {
            window.Show();
            foreach (var locale in new[] { "en", "it", "en" })
            {
                window.DataContext = new { LanguageSelection = new { SelectedLocale = new { Code = locale } } };
                window.UpdateLayout();
                Assert.Equal(locale, window.Language.IetfLanguageTag);
                Assert.Equal(date.ToString("g", window.Language.GetSpecificCulture()), text.Text);
            }
        }
        finally { window.Close(); }
    }

    [Fact]
    public void PresentationProject_UsesWindowsSubsystem()
    {
        var projectPath = TestPathHelper.GetRepositoryPath(
            "src",
            "OnlyExo365.Shell",
            "OnlyExo365.Shell.csproj");

        var project = XDocument.Load(projectPath);

        Assert.Equal(
            "WinExe",
            project.Root?
                .Elements("PropertyGroup")
                .Elements("OutputType")
                .Single()
                .Value);
    }

    [Fact]
    public void WorkerSupervisor_StartsWorkerWithoutVisibleConsoleWindow()
    {
        var supervisorPath = TestPathHelper.GetRepositoryPath(
            "src",
            "OnlyExo365.Shell",
            "Ipc",
            "WorkerSupervisor.cs");

        var source = File.ReadAllText(supervisorPath);

        Assert.Contains("UseShellExecute = false", source, StringComparison.Ordinal);
        Assert.Contains("CreateNoWindow = true", source, StringComparison.Ordinal);
        Assert.Contains("WindowStyle = ProcessWindowStyle.Hidden", source, StringComparison.Ordinal);
        Assert.DoesNotContain("WindowStyle = ProcessWindowStyle.Normal", source, StringComparison.Ordinal);
    }
}

