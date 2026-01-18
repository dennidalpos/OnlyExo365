using System.Windows.Controls;

namespace ExchangeAdmin.Presentation.Views;

/// <summary>
/// Shared mailbox list view (reuses MailboxListViewModel with SharedMailbox filter)
/// </summary>
public partial class SharedMailboxListView : UserControl
{
    public SharedMailboxListView()
    {
        InitializeComponent();
    }
}
