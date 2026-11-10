using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using OnlyExo365.Shell.Localization;
using OnlyExo365.Shell.Services;
using OnlyExo365.Shell.ViewModels;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;
using OnlyExo365.Shell.Ipc;
using OnlyExo365.Shell.Helpers;
using OnlyExo365.Shell.Views;

namespace OnlyExo365.Tests;

public sealed class UiInteractionTests
{
    [WpfFact]
    public async Task Debounce_UpdatesBoundCollectionOnItsDispatcher()
    {
        var items = new ObservableCollection<int> { 1 };
        var view = new ListCollectionView(items);
        var owner = Environment.CurrentManagedThreadId;
        var completion = new TaskCompletionSource<int>();
        using var debounce = new DebounceHelper();
        debounce.Debounce(() =>
        {
            items.Add(2);
            completion.SetResult(Environment.CurrentManagedThreadId);
        }, 10);
        Assert.Equal(owner, await completion.Task.WaitAsync(TimeSpan.FromSeconds(2)));
        Assert.Equal(2, view.Count);
    }

    [WpfFact]
    public async Task Debounce_OnlyRunsLatestCallbackAndCancelsOnDispose()
    {
        var calls = new List<int>();
        var completion = new TaskCompletionSource();
        var debounce = new DebounceHelper();
        debounce.Debounce(() => calls.Add(1), 20);
        debounce.Debounce(() => { calls.Add(2); completion.SetResult(); }, 20);
        await completion.Task.WaitAsync(TimeSpan.FromSeconds(2));
        debounce.Debounce(() => calls.Add(3), 20);
        debounce.Dispose();
        await Task.Delay(80);
        Assert.Equal(new[] { 2 }, calls);
    }

    [WpfFact]
    public void ConfirmationDialog_StartsWithSafeActionAndReturnsFalseOnCancel()
    {
        var dialog = new ConfirmationDialog("Delete", "Delete the selected object")
        {
            Left = -10000, Top = -10000, WindowStartupLocation = WindowStartupLocation.Manual
        };
        dialog.ContentRendered += (_, _) =>
        {
            var cancel = (Button)dialog.FindName("CancelButton");
            Assert.True(cancel.IsCancel);
            Assert.True(cancel.IsDefault);
            Assert.Same(cancel, Keyboard.FocusedElement);
            cancel.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        };
        Assert.False(dialog.ShowDialog());
    }

    [WpfFact]
    public void ConfirmationDialog_ClosingWithoutConfirmationReturnsFalse()
    {
        var dialog = new ConfirmationDialog("Delete", "Target A")
        {
            Left = -10000, Top = -10000, WindowStartupLocation = WindowStartupLocation.Manual
        };
        dialog.ContentRendered += (_, _) => dialog.Close();
        Assert.False(dialog.ShowDialog());
    }

    internal static ResourceDictionary LoadTheme() => (ResourceDictionary)Application.LoadComponent(
        new Uri("/OnlyExo365.Shell;component/Themes/DarkTheme.xaml", UriKind.Relative));

    [WpfFact]
    public void SelectedComboBoxItem_TextUsesTheSelectedForeground()
    {
        var item = new ComboBoxItem { Content = "Selected policy", IsSelected = true, Resources = LoadTheme() };
        item.ApplyTemplate();
        item.Measure(new Size(300, 80));
        item.Arrange(new Rect(0, 0, 300, 80));
        item.UpdateLayout();
        var text = Assert.Single(Descendants(item).OfType<TextBlock>());
        Assert.Equal(Colors.White, ((SolidColorBrush)text.Foreground).Color);
    }

    [WpfFact]
    public void EditableComboBox_ExposesTextAndBindsTypedValue()
    {
        var source = new TextSource { Value = "Delete" };
        var combo = new ComboBox { IsEditable = true, Resources = LoadTheme() };
        combo.Items.Add("Delete");
        combo.SetBinding(ComboBox.TextProperty, new Binding(nameof(TextSource.Value))
        {
            Source = source, Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
        });
        combo.ApplyTemplate();
        var editor = Assert.IsType<TextBox>(combo.Template.FindName("PART_EditableTextBox", combo));
        Assert.Equal(Visibility.Visible, editor.Visibility);
        Assert.NotNull(combo.Template.FindName("PART_Popup", combo));
        editor.Text = "CustomAction";
        Assert.Equal("CustomAction", source.Value);
        combo.SelectedIndex = 0;
        Assert.Equal("Delete", source.Value);
    }

    [WpfFact]
    public void Theme_ActionColorsMeetContrastAndFollowHighContrast()
    {
        var theme = LoadTheme();
        var resources = new ResourceDictionary();
        resources.MergedDictionaries.Add(theme);
        using var palette = new ThemePaletteService(resources);
        palette.Apply(false);
        foreach (var (text, background) in new[]
        {
            ("PrimaryActionTextBrush", "AccentBrush"),
            ("PrimaryActionTextBrush", "AccentHoverBrush"),
            ("DangerActionTextBrush", "DangerBackgroundBrush"),
            ("DangerActionTextBrush", "DangerHoverBackgroundBrush")
        })
        {
            var foreground = ((SolidColorBrush)resources[text]).Color;
            var backdrop = ((SolidColorBrush)resources[background]).Color;
            var first = Luminance(foreground);
            var second = Luminance(backdrop);
            Assert.True((Math.Max(first, second) + 0.05) / (Math.Min(first, second) + 0.05) >= 4.5);
        }
        palette.Apply(true);
        Assert.Same(SystemColors.HighlightBrush, resources["DangerBackgroundBrush"]);
        Assert.Same(SystemColors.HighlightTextBrush, resources["DangerActionTextBrush"]);
        Assert.Same(SystemColors.WindowBrush, resources["PrimaryBackgroundBrush"]);
        palette.Apply(false);
        Assert.Same(theme["AccentBrush"], resources["AccentBrush"]);
    }

    [WpfFact]
    public void DateBinding_ReformatsWhenInheritedLanguageChanges()
    {
        var date = new DateTime(2026, 10, 9, 15, 30, 0);
        var text = new TextBlock();
        text.SetBinding(TextBlock.TextProperty, new Binding { Source = date, StringFormat = "{0:g}" });
        var parent = new StackPanel();
        parent.Children.Add(text);
        foreach (var locale in new[] { "en", "it" })
        {
            parent.Language = XmlLanguage.GetLanguage(locale);
            parent.UpdateLayout();
            Assert.Equal(date.ToString("g", parent.Language.GetSpecificCulture()), text.Text);
        }
    }

    [WpfFact]
    public void BoundedGrid_RealizesViewportRowsInsideOuterScrollViewer()
    {
        var grid = new DataGrid
        {
            MaxHeight = 320, Resources = LoadTheme(), AutoGenerateColumns = false,
            ItemsSource = Enumerable.Range(0, 1000).Select(index => new TextSource { Value = index.ToString() }).ToList()
        };
        grid.Columns.Add(new DataGridTextColumn { Binding = new Binding(nameof(TextSource.Value)) });
        VirtualizingPanel.SetVirtualizationMode(grid, VirtualizationMode.Recycling);
        var window = new Window
        {
            Width = 700, Height = 500, Left = -10000, Top = -10000,
            WindowStartupLocation = WindowStartupLocation.Manual,
            Content = new ScrollViewer { Content = grid }
        };
        try
        {
            window.Show();
            window.UpdateLayout();
            var realized = Enumerable.Range(0, 1000).Count(index => grid.ItemContainerGenerator.ContainerFromIndex(index) != null);
            Assert.InRange(realized, 1, 50);
            Assert.True(grid.ActualHeight <= 320);
        }
        finally { window.Close(); }
    }

    [WpfFact]
    public void TraceLayout_KeepsCommandsAndPagingInsideCompactViewports()
    {
        var originalLocale = LocalizationService.Instance.CurrentLocale;
        try
        {
            foreach (var locale in new[] { "en", "it" })
            foreach (var viewport in new[] { new Size(650, 420), new Size(950, 620) })
            foreach (var withDetails in new[] { false, true })
            {
                LocalizationService.Instance.SetLocale(locale);
                var source = File.ReadAllText(TestPathHelper.GetRepositoryPath(
                    "src", "OnlyExo365.Shell", "Views", "MessageTraceView.xaml"));
                source = source.Replace(" x:Class=\"OnlyExo365.Shell.Views.MessageTraceView\"", string.Empty)
                    .Replace("clr-namespace:OnlyExo365.Shell.Localization", "clr-namespace:OnlyExo365.Shell.Localization;assembly=OnlyExo365.Shell");
                var openingTag = source.IndexOf('>') + 1;
                source = source.Insert(openingTag,
                    "<UserControl.Resources><ResourceDictionary Source=\"pack://application:,,,/OnlyExo365.Shell;component/Themes/DarkTheme.xaml\"/></UserControl.Resources>");
                var view = (UserControl)XamlReader.Parse(source);
                var presentation = new TracePresentation();
                if (withDetails) presentation.SelectedMessageEvents.Add(new MessageTraceDetailEventDto { Event = "Delivered" });
                view.DataContext = new { MessageTrace = presentation };
                view.Language = XmlLanguage.GetLanguage(locale);
                view.Width = viewport.Width;
                view.Height = viewport.Height;
                var host = new Window
                {
                    Content = view, SizeToContent = SizeToContent.WidthAndHeight,
                    Left = -10000, Top = -10000, WindowStartupLocation = WindowStartupLocation.Manual
                };
                try
                {
                    host.Show();
                    host.UpdateLayout();
                    var buttons = Descendants(view).OfType<Button>().Where(button => button.IsVisible).ToList();
                    Assert.NotEmpty(buttons);
                    foreach (var button in buttons)
                    {
                        var bounds = button.TransformToAncestor(view).TransformBounds(new Rect(button.RenderSize));
                        Assert.True(bounds.Right <= viewport.Width + 1 && bounds.Bottom <= viewport.Height + 1,
                            $"{locale} {viewport}: {button.Content} at {bounds}");
                    }
                }
                finally { host.Close(); }
            }
        }
        finally { LocalizationService.Instance.SetLocale(originalLocale); }
    }

    private static IEnumerable<DependencyObject> Descendants(DependencyObject parent)
    {
        for (var index = 0; index < VisualTreeHelper.GetChildrenCount(parent); index++)
        {
            var child = VisualTreeHelper.GetChild(parent, index);
            yield return child;
            foreach (var descendant in Descendants(child)) yield return descendant;
        }
    }

    private sealed class TracePresentation
    {
        public bool HasError => false;
        public bool HasWarnings => false;
        public bool IsLoading => false;
        public ObservableCollection<MessageTraceDetailEventDto> SelectedMessageEvents { get; } = [];
        public ObservableCollection<MessageTraceItemDto> Messages { get; } = [];
    }

    [WpfFact]
    public void PublicFolderDraft_RevertingAFieldClearsDirtyState()
    {
        using var shell = new ShellViewModel(new UiWorker(), new NavigationService());
        var editor = new PublicFoldersViewModel(new UiWorker(), shell);
        Assert.False(editor.HasPendingChanges);
        editor.Name = "Draft";
        Assert.True(editor.HasPendingChanges);
        editor.Name = string.Empty;
        Assert.False(editor.HasPendingChanges);
        editor.ReleaseUiResources();
    }

    [WpfFact]
    public void Navigation_StayingPreservesDraftAndDiscardRunsBeforeLeaving()
    {
        var navigation = new NavigationService();
        var progress = new ShellProgressViewModel();
        var prompt = new ShellPromptViewModel();
        using var state = new ShellNavigationStateViewModel(navigation, progress, prompt);
        var dirty = true;
        state.RegisterUnsavedChangesCheck(() => dirty, () => dirty = false);
        try
        {
            ErrorDialogService.ConfirmationHandlerOverride = (_, _) => false;
            navigation.NavigateTo(NavigationPage.Contacts);
            Assert.Equal(NavigationPage.Dashboard, navigation.CurrentPage);
            Assert.True(dirty);
            ErrorDialogService.ConfirmationHandlerOverride = (_, _) => true;
            navigation.NavigateTo(NavigationPage.Contacts);
            Assert.Equal(NavigationPage.Contacts, navigation.CurrentPage);
            Assert.False(dirty);
        }
        finally
        {
            ErrorDialogService.ConfirmationHandlerOverride = null;
            progress.ReleaseUiResources();
            prompt.ReleaseUiResources();
        }
    }

    private static double Luminance(Color color)
    {
        static double Linear(byte channel)
        {
            var value = channel / 255d;
            return value <= 0.04045 ? value / 12.92 : Math.Pow((value + 0.055) / 1.055, 2.4);
        }
        return 0.2126 * Linear(color.R) + 0.7152 * Linear(color.G) + 0.0722 * Linear(color.B);
    }

    [WpfFact]
    public void MailboxSave_CancellingConfirmationSubmitsNoWrites()
    {
        var worker = new MailboxWriteWorker();
        using var shell = new ShellViewModel(new UiWorker(), new NavigationService());
        var editor = new MailboxDetailsViewModel(worker, new NavigationService(), shell, new CacheService());
        typeof(MailboxDetailsViewModel).GetProperty(nameof(MailboxDetailsViewModel.Identity))!.SetValue(editor, "mailbox-a");
        typeof(MailboxDetailsViewModel).GetProperty(nameof(MailboxDetailsViewModel.Details))!.SetValue(editor,
            new MailboxDetailsDto { Identity = "mailbox-a", PrimarySmtpAddress = "a@contoso.com", Features = new MailboxFeaturesDto() });
        editor.SettingsEditor.HiddenFromAddressListsEnabled = true;
        editor.SettingsEditor.SelectedRetentionPolicy = "policy-a";
        editor.AutoReplyEditor.AutoReplyEnabled = true;
        editor.AutoReplyEditor.AutoReplyInternalMessage = "Away";
        var confirmations = 0;
        ErrorDialogService.ConfirmationHandlerOverride = (_, _) => { confirmations++; return false; };
        try
        {
            Assert.True(editor.SaveMailboxChangesCommand.CanExecute(null));
            editor.SaveMailboxChangesCommand.Execute(null);
            Assert.Equal(1, confirmations);
            Assert.Equal(0, worker.Writes);
            Assert.True(editor.SettingsEditor.HasPendingChanges);
            Assert.True(editor.AutoReplyEditor.HasPendingChanges);
            Assert.True(editor.HasPendingMailboxChanges);
        }
        finally { ErrorDialogService.ConfirmationHandlerOverride = null; editor.ReleaseUiResources(); }
    }

    private sealed class MailboxWriteWorker : TestMailboxesWorkerServiceBase
    {
        public int Writes { get; private set; }
        public override Task<Result> UpdateMailboxSettingsAsync(UpdateMailboxSettingsRequest request,
            Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            Writes++;
            return Task.FromResult(Result.Success());
        }
        public override Task<Result> SetRetentionPolicyAsync(SetRetentionPolicyRequest request,
            Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            Writes++;
            return Task.FromResult(Result.Success());
        }
        public override Task<Result> SetMailboxAutoReplyConfigurationAsync(SetMailboxAutoReplyConfigurationRequest request,
            Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            Writes++;
            return Task.FromResult(Result.Success());
        }
    }

    private sealed class UiWorker : TestWorkerServiceBase { }

    private sealed class TextSource { public string Value { get; set; } = ""; }
}
