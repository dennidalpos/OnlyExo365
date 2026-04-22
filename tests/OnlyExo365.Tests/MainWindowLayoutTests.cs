using System.Xml.Linq;

namespace OnlyExo365.Tests;

public sealed class MainWindowLayoutTests
{
    [Fact]
    public void NavigationSidebar_IsWrappedInVerticalScrollViewer()
    {
        var document = LoadViewDocument("MainWindow.xaml");

        var border = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Border" &&
                string.Equals((string?)candidate.Attribute("Grid.Column"), "0", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Background"), "{StaticResource TertiaryBackgroundBrush}", StringComparison.Ordinal));

        Assert.NotNull(border);

        var scrollViewer = border!
            .Elements()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "ScrollViewer" &&
                string.Equals((string?)candidate.Attribute("VerticalScrollBarVisibility"), "Auto", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("HorizontalScrollBarVisibility"), "Disabled", StringComparison.Ordinal));

        Assert.NotNull(scrollViewer);
    }

    [Fact]
    public void MobileDevicesNavigationButton_UsesCapabilityBindings()
    {
        var document = LoadViewDocument("MainWindow.xaml");

        var button = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Command"), "{Binding NavigateToMobileDevicesCommand}", StringComparison.Ordinal));

        Assert.NotNull(button);
        Assert.Equal("{loc:Loc Key=Nav.MobileDevices}", (string?)button!.Attribute("Content"));
        Assert.Equal("{Binding CanAccessMobileDevicesPage}", (string?)button.Attribute("IsEnabled"));
        Assert.Equal("{Binding MobileDevicesNavigationTooltip}", (string?)button.Attribute("ToolTip"));
    }

    [Fact]
    public void ComplianceNavigationButton_IsPresent()
    {
        var document = LoadViewDocument("MainWindow.xaml");

        var button = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Command"), "{Binding NavigateToComplianceCommand}", StringComparison.Ordinal));

        Assert.NotNull(button);
        Assert.Equal("{loc:Loc Key=Nav.Compliance}", (string?)button!.Attribute("Content"));
        Assert.Equal("{Binding NavigateToComplianceCommand}", (string?)button.Attribute("Command"));
        Assert.Equal("{Binding CanAccessCompliancePage}", (string?)button.Attribute("IsEnabled"));
        Assert.Equal("{Binding ComplianceNavigationTooltip}", (string?)button.Attribute("ToolTip"));
    }

    [Fact]
    public void MailSecurityNavigationButton_IsPresent()
    {
        var document = LoadViewDocument("MainWindow.xaml");

        var button = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Command"), "{Binding NavigateToMailSecurityCommand}", StringComparison.Ordinal));

        Assert.NotNull(button);
        Assert.Equal("{loc:Loc Key=Nav.MailSecurity}", (string?)button!.Attribute("Content"));
        Assert.Equal("{Binding NavigateToMailSecurityCommand}", (string?)button.Attribute("Command"));
        Assert.Equal("{Binding CanAccessMailSecurityPage}", (string?)button.Attribute("IsEnabled"));
        Assert.Equal("{Binding MailSecurityNavigationTooltip}", (string?)button.Attribute("ToolTip"));
    }

    [Fact]
    public void MigrationNavigationButton_UsesLeastPrivilegeBindings()
    {
        var document = LoadViewDocument("MainWindow.xaml");

        var button = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Command"), "{Binding NavigateToMigrationCommand}", StringComparison.Ordinal));

        Assert.NotNull(button);
        Assert.Equal("{loc:Loc Key=Nav.Migration}", (string?)button!.Attribute("Content"));
        Assert.Equal("{Binding CanAccessMigrationPage}", (string?)button.Attribute("IsEnabled"));
        Assert.Equal("{Binding MigrationNavigationTooltip}", (string?)button.Attribute("ToolTip"));
    }

    [Fact]
    public void PermissionsNavigationButton_UsesLeastPrivilegeBindings()
    {
        var document = LoadViewDocument("MainWindow.xaml");

        var button = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Command"), "{Binding NavigateToPermissionsCommand}", StringComparison.Ordinal));

        Assert.NotNull(button);
        Assert.Equal("{loc:Loc Key=Nav.RoleGroups}", (string?)button!.Attribute("Content"));
        Assert.Equal("{Binding CanAccessPermissionsPage}", (string?)button.Attribute("IsEnabled"));
        Assert.Equal("{Binding PermissionsNavigationTooltip}", (string?)button.Attribute("ToolTip"));
    }

    [Fact]
    public void MessageTraceNavigationButton_UsesLeastPrivilegeBindings()
    {
        var document = LoadViewDocument("MainWindow.xaml");

        var button = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Command"), "{Binding NavigateToMessageTraceCommand}", StringComparison.Ordinal));

        Assert.NotNull(button);
        Assert.Equal("{loc:Loc Key=Nav.MessageTrace}", (string?)button!.Attribute("Content"));
        Assert.Equal("{Binding CanAccessMessageTracePage}", (string?)button.Attribute("IsEnabled"));
        Assert.Equal("{Binding MessageTraceNavigationTooltip}", (string?)button.Attribute("ToolTip"));
    }

    [Fact]
    public void GraphStatusText_UsesTooltipBinding()
    {
        var document = LoadViewDocument("MainWindow.xaml");

        var textBlock = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "TextBlock" &&
                string.Equals((string?)candidate.Attribute("Text"), "{Binding GraphStateDisplay}", StringComparison.Ordinal));

        Assert.NotNull(textBlock);
        Assert.Equal("{Binding GraphStateTooltip}", (string?)textBlock!.Attribute("ToolTip"));
    }

    [Fact]
    public void HeaderLayout_SeparatesStatusIndicatorsFromCommandButtons()
    {
        var document = LoadViewDocument("MainWindow.xaml");

        var headerGrid = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Grid" &&
                candidate.Elements().Count(element => element.Name.LocalName == "Grid.RowDefinitions") == 1 &&
                candidate.Elements().Any(element =>
                    element.Name.LocalName == "WrapPanel" &&
                    string.Equals((string?)element.Attribute("Grid.Row"), "1", StringComparison.Ordinal)) &&
                candidate.Elements().Any(element =>
                    element.Name.LocalName == "WrapPanel" &&
                    string.Equals((string?)element.Attribute("Grid.Row"), "2", StringComparison.Ordinal)));

        Assert.NotNull(headerGrid);

        var statusPanel = headerGrid!
            .Elements()
            .First(candidate =>
                candidate.Name.LocalName == "WrapPanel" &&
                string.Equals((string?)candidate.Attribute("Grid.Row"), "2", StringComparison.Ordinal));

        Assert.Contains(statusPanel.Descendants(), candidate => candidate.Name.LocalName == "TextBlock" && (string?)candidate.Attribute("Text") == "{Binding WorkerStateDisplay}");
        Assert.Contains(statusPanel.Descendants(), candidate => candidate.Name.LocalName == "TextBlock" && (string?)candidate.Attribute("Text") == "{Binding ExchangeStateDisplay}");
        Assert.Contains(statusPanel.Descendants(), candidate => candidate.Name.LocalName == "TextBlock" && (string?)candidate.Attribute("Text") == "{Binding GraphStateDisplay}");

        var commandPanel = headerGrid
            .Elements()
            .First(candidate =>
                candidate.Name.LocalName == "WrapPanel" &&
                string.Equals((string?)candidate.Attribute("Grid.Row"), "1", StringComparison.Ordinal));

        Assert.Contains(commandPanel.Elements(), candidate => candidate.Name.LocalName == "Button" && (string?)candidate.Attribute("Content") == "{loc:Loc Key=Btn.StartWorker}");
        Assert.Contains(commandPanel.Elements(), candidate => candidate.Name.LocalName == "Button" && (string?)candidate.Attribute("Content") == "{loc:Loc Key=Btn.ConnectExchange}");
        Assert.Contains(commandPanel.Elements(), candidate => candidate.Name.LocalName == "Button" && (string?)candidate.Attribute("Content") == "{loc:Loc Key=Btn.Disconnect}");
    }

    [Fact]
    public void Footer_ExposesGlobalAuditTaskCenterBindings()
    {
        var document = LoadViewDocument("MainWindow.xaml");

        var comboBox = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "ComboBox" &&
                string.Equals((string?)candidate.Attribute("ItemsSource"), "{Binding Compliance.AuditSearchTasks}", StringComparison.Ordinal));

        Assert.NotNull(comboBox);
        Assert.Equal("{Binding Compliance.SelectedAuditSearchTask}", (string?)comboBox!.Attribute("SelectedItem"));
        Assert.Equal("TaskSummary", (string?)comboBox.Attribute("DisplayMemberPath"));
    }

    [Fact]
    public void MainContent_ExposesInlineShellPromptOverlay()
    {
        var document = LoadViewDocument("MainWindow.xaml");

        var promptBorder = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Border" &&
                string.Equals((string?)candidate.Attribute("Panel.ZIndex"), "100", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Visibility"), "{Binding Prompt.IsOpen, Converter={StaticResource BoolToVisibility}}", StringComparison.Ordinal));

        Assert.NotNull(promptBorder);

        var confirmButton = promptBorder!
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Content"), "{loc:Loc Key=Btn.Confirm}", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Command"), "{Binding Prompt.ConfirmCommand}", StringComparison.Ordinal));

        Assert.NotNull(confirmButton);
    }

    [Fact]
    public void MainWindow_ExposesSingleGlobalAlertHost_BetweenHeaderAndContent()
    {
        var document = LoadViewDocument("MainWindow.xaml");

        var alertBorder = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Border" &&
                string.Equals((string?)candidate.Attribute("Grid.Row"), "1", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Visibility"), "{Binding GlobalAlert.IsVisible, Converter={StaticResource BoolToVisibility}}", StringComparison.Ordinal));

        Assert.NotNull(alertBorder);
        Assert.Contains(alertBorder!.Descendants(), candidate =>
            candidate.Name.LocalName == "TextBlock" &&
            string.Equals((string?)candidate.Attribute("Text"), "{Binding GlobalAlert.Title}", StringComparison.Ordinal));
        Assert.Contains(alertBorder.Descendants(), candidate =>
            candidate.Name.LocalName == "TextBlock" &&
            string.Equals((string?)candidate.Attribute("Text"), "{Binding GlobalAlert.Message}", StringComparison.Ordinal));
    }

    [Fact]
    public void NavigationSidebar_DoesNotExposeSeparateSharedMailboxesButton()
    {
        var document = LoadViewDocument("MainWindow.xaml");

        var button = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Content"), "Shared Mailboxes", StringComparison.Ordinal));

        Assert.Null(button);
    }

    [Fact]
    public void DarkTheme_DefinesPrimaryBrushAlias()
    {
        var document = LoadThemeDocument("DarkTheme.xaml");

        var brush = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "SolidColorBrush" &&
                string.Equals((string?)candidate.Attribute(XName.Get("Key", "http://schemas.microsoft.com/winfx/2006/xaml")), "PrimaryBrush", StringComparison.Ordinal));

        Assert.NotNull(brush);
        Assert.Equal("{StaticResource AccentColor}", (string?)brush!.Attribute("Color"));
    }

    [Fact]
    public void DarkTheme_ButtonStyle_DefinesConsistentMinimumSize()
    {
        var document = LoadThemeDocument("DarkTheme.xaml");

        var style = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Style" &&
                string.Equals((string?)candidate.Attribute("TargetType"), "Button", StringComparison.Ordinal));

        Assert.NotNull(style);
        Assert.Contains(style!.Elements(), element =>
            element.Name.LocalName == "Setter" &&
            string.Equals((string?)element.Attribute("Property"), "MinWidth", StringComparison.Ordinal) &&
            string.Equals((string?)element.Attribute("Value"), "96", StringComparison.Ordinal));
        Assert.Contains(style.Elements(), element =>
            element.Name.LocalName == "Setter" &&
            string.Equals((string?)element.Attribute("Property"), "MinHeight", StringComparison.Ordinal) &&
            string.Equals((string?)element.Attribute("Value"), "32", StringComparison.Ordinal));
    }

    [Fact]
    public void DarkTheme_ButtonStyle_ShowsDisabledTooltipsAndKeyboardFocus()
    {
        var document = LoadThemeDocument("DarkTheme.xaml");

        var style = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Style" &&
                string.Equals((string?)candidate.Attribute("TargetType"), "Button", StringComparison.Ordinal));

        Assert.NotNull(style);
        Assert.Contains(style!.Elements(), element =>
            element.Name.LocalName == "Setter" &&
            string.Equals((string?)element.Attribute("Property"), "ToolTipService.ShowOnDisabled", StringComparison.Ordinal) &&
            string.Equals((string?)element.Attribute("Value"), "True", StringComparison.Ordinal));

        var keyboardFocusTrigger = style
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Trigger" &&
                string.Equals((string?)candidate.Attribute("Property"), "IsKeyboardFocused", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Value"), "True", StringComparison.Ordinal));

        Assert.NotNull(keyboardFocusTrigger);
        Assert.Contains(keyboardFocusTrigger!.Elements(), element =>
            element.Name.LocalName == "Setter" &&
            string.Equals((string?)element.Attribute("Property"), "BorderBrush", StringComparison.Ordinal) &&
            string.Equals((string?)element.Attribute("Value"), "{StaticResource AccentBrush}", StringComparison.Ordinal));
    }

    [Fact]
    public void DarkTheme_ButtonTemplate_RespectsContentAlignment()
    {
        var document = LoadThemeDocument("DarkTheme.xaml");

        var style = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Style" &&
                string.Equals((string?)candidate.Attribute("TargetType"), "Button", StringComparison.Ordinal));

        Assert.NotNull(style);

        var contentPresenter = style!
            .Descendants()
            .FirstOrDefault(candidate => candidate.Name.LocalName == "ContentPresenter");

        Assert.NotNull(contentPresenter);
        Assert.Equal("{TemplateBinding HorizontalContentAlignment}", (string?)contentPresenter!.Attribute("HorizontalAlignment"));
        Assert.Equal("{TemplateBinding VerticalContentAlignment}", (string?)contentPresenter.Attribute("VerticalAlignment"));
    }

    [Fact]
    public void DarkTheme_ComboBoxStyle_ExposesKeyboardFocusBorder()
    {
        var document = LoadThemeDocument("DarkTheme.xaml");

        var style = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Style" &&
                string.Equals((string?)candidate.Attribute("TargetType"), "ComboBox", StringComparison.Ordinal));

        Assert.NotNull(style);

        var toggleButton = style!
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "ToggleButton" &&
                string.Equals((string?)candidate.Attribute(XName.Get("Name", "http://schemas.microsoft.com/winfx/2006/xaml")), "ToggleButton", StringComparison.Ordinal));

        Assert.NotNull(toggleButton);
        Assert.Equal("{TemplateBinding BorderBrush}", (string?)toggleButton!.Attribute("BorderBrush"));
        Assert.Equal("{TemplateBinding BorderThickness}", (string?)toggleButton.Attribute("BorderThickness"));

        var keyboardFocusTrigger = style
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Trigger" &&
                string.Equals((string?)candidate.Attribute("Property"), "IsKeyboardFocusWithin", StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Value"), "True", StringComparison.Ordinal));

        Assert.NotNull(keyboardFocusTrigger);
        Assert.Contains(keyboardFocusTrigger!.Elements(), element =>
            element.Name.LocalName == "Setter" &&
            string.Equals((string?)element.Attribute("Property"), "BorderBrush", StringComparison.Ordinal) &&
            string.Equals((string?)element.Attribute("Value"), "{StaticResource AccentBrush}", StringComparison.Ordinal));
    }

    [Fact]
    public void DarkTheme_DefinesSharedActionButtonStyles()
    {
        var document = LoadThemeDocument("DarkTheme.xaml");
        var styleKeys = document
            .Descendants()
            .Where(candidate => candidate.Name.LocalName == "Style")
            .Select(candidate => (string?)candidate.Attribute(XName.Get("Key", "http://schemas.microsoft.com/winfx/2006/xaml")))
            .Where(static key => !string.IsNullOrWhiteSpace(key))
            .ToList();

        Assert.Contains("HeaderCommandButton", styleKeys);
        Assert.Contains("HeaderPrimaryCommandButton", styleKeys);
        Assert.Contains("InlineActionButton", styleKeys);
        Assert.Contains("InlinePrimaryActionButton", styleKeys);
        Assert.Contains("InlineDangerActionButton", styleKeys);
        Assert.Contains("ToolbarButton", styleKeys);
        Assert.Contains("ToolbarPrimaryButton", styleKeys);
        Assert.Contains("ToolbarDangerButton", styleKeys);
    }

    [Fact]
    public void NavigationButtons_DeclareActiveSectionHighlightStyle()
    {
        var document = LoadViewDocument("MainWindow.xaml");

        var buttons = document
            .Descendants()
            .Where(candidate => candidate.Name.LocalName == "Button")
            .Where(candidate => candidate.Attribute("Command")?.Value?.Contains("NavigateTo", StringComparison.Ordinal) == true)
            .Where(candidate => candidate.Elements().Any(descendant => descendant.Name.LocalName == "Button.Style"))
            .ToList();

        Assert.NotEmpty(buttons);

        foreach (var button in buttons)
        {
            var style = button
                .Elements()
                .FirstOrDefault(candidate => candidate.Name.LocalName == "Button.Style")
                ?.Elements()
                .FirstOrDefault(candidate =>
                    candidate.Name.LocalName == "Style" &&
                    string.Equals((string?)candidate.Attribute("BasedOn"), "{StaticResource NavButton}", StringComparison.Ordinal));

            Assert.NotNull(style);

            var accentTrigger = style!
                .Descendants()
                .FirstOrDefault(candidate =>
                    candidate.Name.LocalName == "DataTrigger" &&
                    string.Equals((string?)candidate.Attribute("Value"), "True", StringComparison.Ordinal) &&
                    candidate.Elements().Any(descendant =>
                        descendant.Name.LocalName == "Setter" &&
                        string.Equals((string?)descendant.Attribute("Property"), "Background", StringComparison.Ordinal) &&
                        string.Equals((string?)descendant.Attribute("Value"), "{StaticResource AccentBrush}", StringComparison.Ordinal)));

            Assert.NotNull(accentTrigger);
        }
    }

    [Fact]
    public void MainWindow_HeaderCommands_UseSharedHeaderButtonStyles()
    {
        var document = LoadViewDocument("MainWindow.xaml");

        AssertButtonStyle(document, "{loc:Loc Key=Btn.StartWorker}", "{StaticResource HeaderPrimaryCommandButton}");
        AssertButtonStyle(document, "{loc:Loc Key=Btn.StopWorker}", "{StaticResource HeaderCommandButton}");
        AssertButtonStyle(document, "{loc:Loc Key=Btn.ConnectExchange}", "{StaticResource HeaderPrimaryCommandButton}");
        AssertButtonStyle(document, "{loc:Loc Key=Btn.Disconnect}", "{StaticResource HeaderCommandButton}");
    }

    [Fact]
    public void MailboxLicensesTab_UsesSharedInlineActionStyles()
    {
        var document = LoadViewDocument("MailboxLicensesTab.xaml");

        AssertButtonStyle(document, "{loc:Loc Key=Btn.Remove}", "{StaticResource InlineDangerActionButton}");
        AssertButtonStyle(document, "{loc:Loc Key=Btn.Assign}", "{StaticResource InlinePrimaryActionButton}");
        AssertButtonStyle(document, "{loc:Loc Key=Btn.Refresh}", "{StaticResource InlineActionButton}");
    }

    [Fact]
    public void MigrationView_UsesSharedToolbarButtonStyles()
    {
        var document = LoadViewDocument("MigrationView.xaml");

        AssertButtonStyle(document, "{loc:Loc Key=Migration.RefreshBatches}", "{StaticResource ToolbarPrimaryButton}");
        AssertButtonStyle(document, "{loc:Loc Key=Migration.RefreshEndpoints}", "{StaticResource ToolbarPrimaryButton}");
        AssertButtonStyle(document, "{loc:Loc Key=Migration.SaveEndpoint}", "{StaticResource ToolbarPrimaryButton}");
        AssertButtonStyle(document, "{loc:Loc Key=Migration.Remove}", "{StaticResource ToolbarDangerButton}");
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

