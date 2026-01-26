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

    
    
    

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        
        
        
        

        
        var workerOptions = new WorkerSupervisorOptions
        {
            WorkerPath = "ExchangeAdmin.Worker.exe",
            MaxRestartAttempts = 3,
            ExchangeEnvironmentName = "O365Default"
        };

        _workerService = new WorkerService(workerOptions);
        _navigationService = new NavigationService();

        
        _shellViewModel = new ShellViewModel(_workerService, _navigationService);

        
        var dashboardViewModel = new DashboardViewModel(_workerService, _navigationService, _shellViewModel);
        var mailboxListViewModel = new MailboxListViewModel(_workerService, _navigationService, _shellViewModel);
        var sharedMailboxListViewModel = new MailboxListViewModel(_workerService, _navigationService, _shellViewModel);
        var mailboxDetailsViewModel = new MailboxDetailsViewModel(_workerService, _navigationService, _shellViewModel);
        var mailboxSpaceViewModel = new MailboxSpaceViewModel(_workerService, _shellViewModel);
        var distributionListViewModel = new DistributionListViewModel(_workerService, _navigationService, _shellViewModel);
        var logsViewModel = new LogsViewModel(_shellViewModel);

        
        sharedMailboxListViewModel.RecipientTypeFilter = "SharedMailbox";

        
        
        _shellViewModel.Dashboard = dashboardViewModel;
        _shellViewModel.Mailboxes = mailboxListViewModel;
        _shellViewModel.SharedMailboxes = sharedMailboxListViewModel;
        _shellViewModel.MailboxDetails = mailboxDetailsViewModel;
        _shellViewModel.MailboxSpace = mailboxSpaceViewModel;
        _shellViewModel.DistributionLists = distributionListViewModel;
        _shellViewModel.Logs = logsViewModel;

        
        var mainWindow = new Views.MainWindow
        {
            DataContext = _shellViewModel
        };

        

        
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

        
        _shellViewModel.PropertyChanged += async (s, e) =>
        {
            if (e.PropertyName == nameof(ShellViewModel.IsExchangeConnected) && _shellViewModel.IsExchangeConnected)
            {
                
                
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
        

        try
        {
            _shellViewModel?.Dispose();

            if (_workerService is IAsyncDisposable disposable)
            {
                
                
                
                disposable.DisposeAsync().AsTask().GetAwaiter().GetResult();
            }

            
        }
        catch
        {
            
        }

        base.OnExit(e);
    }
}
