using System.Windows.Controls;
using ExchangeAdmin.Presentation.ViewModels;

namespace ExchangeAdmin.Presentation.Views;

public partial class ToolsView : UserControl
{
    public ToolsView()
    {
        InitializeComponent();
    }

    private void WorkerConsoleToggle_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        if (sender is not CheckBox checkBox)
        {
            return;
        }

        var requestedVisibility = checkBox.IsChecked == true;
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
