using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using ExchangeAdmin.Presentation.ViewModels;

namespace ExchangeAdmin.Presentation.Views;

public partial class ToolsView : UserControl
{
    public ToolsView()
    {
        InitializeComponent();
    }

    private void WorkerConsoleToggle_Changed(object sender, System.Windows.RoutedEventArgs e)
    {
        if (sender is not CheckBox)
        {
            return;
        }

        bool? requestedVisibility = e.RoutedEvent switch
        {
            var routedEvent when routedEvent == ToggleButton.CheckedEvent => true,
            var routedEvent when routedEvent == ToggleButton.UncheckedEvent => false,
            _ => null
        };

        var command = DataContext switch
        {
            ShellViewModel shellViewModel => shellViewModel.SetWorkerConsoleVisibilityCommand,
            ToolsViewModel toolsViewModel => toolsViewModel.SetWorkerConsoleVisibilityCommand,
            _ => null
        };

        if (command == null || !command.CanExecute(requestedVisibility))
        {
            return;
        }

        command.Execute(requestedVisibility);
    }
}
