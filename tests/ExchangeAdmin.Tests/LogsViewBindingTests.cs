using System.Xml.Linq;

namespace ExchangeAdmin.Tests;

public sealed class LogsViewBindingTests
{
    [Fact]
    public void LogsView_ExposesPersistentLogPolicyAndExportBindings()
    {
        var document = LoadViewDocument("LogsView.xaml");

        var exportButton = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Content"), "{loc:Loc Key=Btn.ExportBundle}", StringComparison.Ordinal));

        Assert.NotNull(exportButton);
        Assert.Equal("{Binding Logs.ExportPersistentLogsCommand}", (string?)exportButton!.Attribute("Command"));

        var retentionText = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "TextBlock" &&
                string.Equals((string?)candidate.Attribute("Text"), "{Binding Logs.PersistentLogPolicySummary}", StringComparison.Ordinal));

        Assert.NotNull(retentionText);

        var totalCountText = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "TextBlock" &&
                string.Equals((string?)candidate.Attribute("Text"), "{Binding Logs.TotalCount, StringFormat={loc:Loc Key=Logs.EntryCount}}", StringComparison.Ordinal));

        Assert.NotNull(totalCountText);

        var directoryText = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "TextBlock" &&
                string.Equals((string?)candidate.Attribute("Text"), "{Binding Logs.PersistentLogDirectory}", StringComparison.Ordinal));

        Assert.NotNull(directoryText);
        Assert.Equal("{Binding Logs.PersistentLogDirectory}", (string?)directoryText!.Attribute("ToolTip"));

        var summaryPathText = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "TextBlock" &&
                string.Equals((string?)candidate.Attribute("Text"), "{Binding Logs.PersistentObservabilitySummaryPath}", StringComparison.Ordinal));

        Assert.NotNull(summaryPathText);
        Assert.Equal("{Binding Logs.PersistentObservabilitySummaryPath}", (string?)summaryPathText!.Attribute("ToolTip"));
    }

    [Fact]
    public void LogsView_ExposesVisibleLabelsAndAssistiveNamesForFilters()
    {
        var document = LoadViewDocument("LogsView.xaml");

        var searchLabel = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "TextBlock" &&
                string.Equals((string?)candidate.Attribute("Text"), "{loc:Loc Key=Logs.SearchLabel}", StringComparison.Ordinal));

        Assert.NotNull(searchLabel);

        var searchBox = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "TextBox" &&
                string.Equals((string?)candidate.Attribute("Text"), "{Binding Logs.SearchFilter, UpdateSourceTrigger=PropertyChanged}", StringComparison.Ordinal));

        Assert.NotNull(searchBox);
        Assert.Equal("{loc:Loc Key=Logs.SearchTooltip}", (string?)searchBox!.Attribute("ToolTip"));
        Assert.Equal("{loc:Loc Key=Logs.SearchLabel}", (string?)searchBox.Attribute("AutomationProperties.Name"));

        var levelLabel = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "TextBlock" &&
                string.Equals((string?)candidate.Attribute("Text"), "{loc:Loc Key=Logs.LevelFilterLabel}", StringComparison.Ordinal));

        Assert.NotNull(levelLabel);

        var levelCombo = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "ComboBox" &&
                string.Equals((string?)candidate.Attribute("SelectedValue"), "{Binding Logs.FilterLevel, Mode=TwoWay}", StringComparison.Ordinal));

        Assert.NotNull(levelCombo);
        Assert.Equal("{loc:Loc Key=Logs.FilterTooltip}", (string?)levelCombo!.Attribute("ToolTip"));
        Assert.Equal("{loc:Loc Key=Logs.LevelFilterLabel}", (string?)levelCombo.Attribute("AutomationProperties.Name"));
    }

    private static XDocument LoadViewDocument(string viewFileName)
    {
        var viewPath = TestPathHelper.GetRepositoryPath("src", "ExchangeAdmin.Presentation", "Views", viewFileName);

        return XDocument.Load(viewPath);
    }
}
