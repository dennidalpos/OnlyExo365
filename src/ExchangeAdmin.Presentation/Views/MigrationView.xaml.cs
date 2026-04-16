using System.Windows.Controls;
using ExchangeAdmin.Presentation.ViewModels;

namespace ExchangeAdmin.Presentation.Views;

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
