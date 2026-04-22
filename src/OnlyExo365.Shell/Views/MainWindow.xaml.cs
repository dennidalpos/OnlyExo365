using System.ComponentModel;
using System.Windows;

namespace OnlyExo365.Shell.Views;

public partial class MainWindow : Window
{
    private bool _allowClose;
    private Task? _shutdownTask;

    public MainWindow()
    {
        InitializeComponent();
    }

    public Func<Task>? ShutdownHandler { get; set; }

    protected override void OnClosing(CancelEventArgs e)
    {
        if (_allowClose || ShutdownHandler is null)
        {
            base.OnClosing(e);
            return;
        }

        e.Cancel = true;
        IsEnabled = false;
        _shutdownTask ??= ShutdownAndCloseAsync();

        base.OnClosing(e);
    }

    protected override void OnClosed(EventArgs e)
    {
        if (DataContext is IDisposable disposable)
        {
            disposable.Dispose();
        }

        base.OnClosed(e);
    }

    private async Task ShutdownAndCloseAsync()
    {
        try
        {
            await ShutdownHandler!();
        }
        finally
        {
            await Dispatcher.InvokeAsync(() =>
            {
                IsEnabled = true;
                _allowClose = true;
                Close();
            });
        }
    }
}

