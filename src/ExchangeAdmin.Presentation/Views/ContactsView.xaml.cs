using System.Windows.Controls;
using ExchangeAdmin.Presentation.ViewModels;

namespace ExchangeAdmin.Presentation.Views;

public partial class ContactsView : UserControl
{
    public ContactsView()
    {
        InitializeComponent();
    }

    private void MailUserPasswordBox_OnPasswordChanged(object sender, System.Windows.RoutedEventArgs e)
    {
        if (DataContext is ShellViewModel { Contacts: not null } shell &&
            sender is PasswordBox passwordBox)
        {
            shell.Contacts.SetMailUserPassword(passwordBox.Password);
        }
    }
}
