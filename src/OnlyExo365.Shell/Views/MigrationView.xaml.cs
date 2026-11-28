using System.Windows.Controls;
using OnlyExo365.Shell.ViewModels;

namespace OnlyExo365.Shell.Views;

public partial class MigrationView : UserControl
{
    public MigrationView()
    {
        InitializeComponent();
    }

    private void EndpointPasswordBox_OnPasswordChanged(object sender, System.Windows.RoutedEventArgs e)
    {
        if (DataContext is ShellViewModel { Migration: not null } shell &&
            sender is PasswordBox passwordBox)
        {
            shell.Migration.SetEndpointPassword(passwordBox.Password);
        }
    }
}

