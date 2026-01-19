using System.Runtime.InteropServices;
using System.Windows;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Infrastructure.Ipc;
using ExchangeAdmin.Presentation.Services;
using ExchangeAdmin.Presentation.ViewModels;

namespace ExchangeAdmin.Presentation;

public partial class App : System.Windows.Application
{
    private IWorkerService? _workerService;
    private NavigationService? _navigationService;
    private ShellViewModel? _shellViewModel;

    // [DllImport("kernel32.dll", SetLastError = true)]
    // [return: MarshalAs(UnmanagedType.Bool)]
    // private static extern bool AllocConsole();

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Console disabled - all logs now visible in UI Logs tab
        // AllocConsole();
        // Console.WriteLine("[App] Console allocated - UI logs will appear here");
        // Console.WriteLine("[App] Starting ExchangeAdmin Presentation...");

        // Configure services
        var workerOptions = new WorkerSupervisorOptions
        {
            WorkerPath = "ExchangeAdmin.Worker.exe",
            MaxRestartAttempts = 3
        };

        _workerService = new WorkerService(workerOptions);
        _navigationService = new NavigationService();

        // Create Shell ViewModel
        _shellViewModel = new ShellViewModel(_workerService, _navigationService);

        // Create child ViewModels
        var dashboardViewModel = new DashboardViewModel(_workerService, _navigationService, _shellViewModel);
        var mailboxListViewModel = new MailboxListViewModel(_workerService, _navigationService, _shellViewModel);
        var sharedMailboxListViewModel = new MailboxListViewModel(_workerService, _navigationService, _shellViewModel);
        var mailboxDetailsViewModel = new MailboxDetailsViewModel(_workerService, _navigationService, _shellViewModel);
        var distributionListViewModel = new DistributionListViewModel(_workerService, _navigationService, _shellViewModel);
        var logsViewModel = new LogsViewModel(_shellViewModel);

        // Set default filter for shared mailbox view
        sharedMailboxListViewModel.RecipientTypeFilter = "SharedMailbox";

        // Expose child ViewModels through ShellViewModel
        // Console.WriteLine("[App] Connecting child ViewModels to Shell...");
        _shellViewModel.Dashboard = dashboardViewModel;
        _shellViewModel.Mailboxes = mailboxListViewModel;
        _shellViewModel.SharedMailboxes = sharedMailboxListViewModel;
        _shellViewModel.MailboxDetails = mailboxDetailsViewModel;
        _shellViewModel.DistributionLists = distributionListViewModel;
        _shellViewModel.Logs = logsViewModel;

        // Create MainWindow with Shell as DataContext (ALL views use ShellViewModel)
        var mainWindow = new Views.MainWindow
        {
            DataContext = _shellViewModel
        };

        // Console.WriteLine("[App] MainWindow DataContext set to ShellViewModel (child VMs exposed through Shell)");

        // Listen for navigation changes to load data
        _navigationService.PageChanged += async (s, page) =>
        {
            switch (page)
            {
                case NavigationPage.Dashboard:
                    await dashboardViewModel.LoadAsync();
                    break;
                case NavigationPage.Mailboxes:
                    mailboxListViewModel.RecipientTypeFilter = "UserMailbox";
                    await mailboxListViewModel.LoadAsync();
                    break;
                case NavigationPage.SharedMailboxes:
                    sharedMailboxListViewModel.RecipientTypeFilter = "SharedMailbox";
                    await sharedMailboxListViewModel.LoadAsync();
                    break;
                case NavigationPage.DistributionLists:
                    await distributionListViewModel.LoadAsync();
                    break;
                case NavigationPage.Logs:
                    logsViewModel.Refresh();
                    break;
            }
        };

        // Listen for Exchange connection changes to reload current page
        _shellViewModel.PropertyChanged += async (s, e) =>
        {
            if (e.PropertyName == nameof(ShellViewModel.IsExchangeConnected) && _shellViewModel.IsExchangeConnected)
            {
                // Console.WriteLine("[App] Exchange connected - reloading current page data");
                // Reload data for the current page
                switch (_navigationService.CurrentPage)
                {
                    case NavigationPage.Dashboard:
                        await dashboardViewModel.LoadAsync();
                        break;
                    case NavigationPage.Mailboxes:
                        await mailboxListViewModel.LoadAsync();
                        break;
                    case NavigationPage.SharedMailboxes:
                        await sharedMailboxListViewModel.LoadAsync();
                        break;
                    case NavigationPage.DistributionLists:
                        await distributionListViewModel.LoadAsync();
                        break;
                }
            }
        };

        MainWindow = mainWindow;
        mainWindow.Show();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        // Console.WriteLine("[App] Application exiting - starting cleanup...");

        try
        {
            _shellViewModel?.Dispose();

            if (_workerService is IAsyncDisposable disposable)
            {
                // CRITICAL: Use GetAwaiter().GetResult() to wait synchronously
                // async void in OnExit causes exit code -1073741510 (STATUS_CONTROL_C_EXIT)
                // because the process terminates before async cleanup completes
                disposable.DisposeAsync().AsTask().GetAwaiter().GetResult();
            }

            // Console.WriteLine("[App] Cleanup completed successfully");
        }
        catch
        {
            // Console.WriteLine($"[App] Error during cleanup: {ex.Message}");
        }

        base.OnExit(e);
    }
}
