using System.Windows.Controls;
using ExchangeAdmin.Presentation.ViewModels;

namespace ExchangeAdmin.Presentation.Views;

public partial class MailboxListView : UserControl
{
    public MailboxListView()
    {
        InitializeComponent();
    }

    private void NewMailboxPasswordBox_OnPasswordChanged(object sender, System.Windows.RoutedEventArgs e)
    {
        if (DataContext is ShellViewModel { Mailboxes: not null } shell &&
            sender is PasswordBox passwordBox)
        {
            shell.Mailboxes.SetNewMailboxPassword(passwordBox.Password);
        }
    }
}
