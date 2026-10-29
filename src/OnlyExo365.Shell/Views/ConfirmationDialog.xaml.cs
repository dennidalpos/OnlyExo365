using System.Windows;
using OnlyExo365.Shell.Text;

namespace OnlyExo365.Shell.Views;

public partial class ConfirmationDialog : Window
{
    public ConfirmationDialog(string title, string message, string? confirmLabel = null, string? cancelLabel = null)
    {
        InitializeComponent();
        Title = title;
        MessageText.Text = message;
        ConfirmButton.Content = confirmLabel ?? UiTextCatalog.ConfirmButton;
        CancelButton.Content = cancelLabel ?? UiTextCatalog.CancelButton;
        Loaded += (_, _) => CancelButton.Focus();
    }

    private void OnConfirm(object sender, RoutedEventArgs e) => DialogResult = true;
    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;
}
