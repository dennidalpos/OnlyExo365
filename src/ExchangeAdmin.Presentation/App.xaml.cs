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
    private CacheService? _cacheService;
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
        _cacheService = new CacheService();

        // Create shell view model
        _shellViewModel = new ShellViewModel(_workerService, _navigationService);

        // Create child view models with cache service
        var dashboardViewModel = new DashboardViewModel(_workerService, _navigationService, _shellViewModel, _cacheService);
        var mailboxListViewModel = new MailboxListViewModel(_workerService, _navigationService, _shellViewModel);
        var deletedMailboxViewModel = new DeletedMailboxesViewModel(_workerService, _shellViewModel);
        var sharedMailboxListViewModel = new MailboxListViewModel(_workerService, _navigationService, _shellViewModel);
        var mailboxDetailsViewModel = new MailboxDetailsViewModel(_workerService, _navigationService, _shellViewModel, _cacheService);
        var mailboxSpaceViewModel = new MailboxSpaceViewModel(_workerService, _navigationService, _shellViewModel);
        var distributionListViewModel = new DistributionListViewModel(_workerService, _navigationService, _shellViewModel);
        var logsViewModel = new LogsViewModel(_shellViewModel);
        var toolsViewModel = new ToolsViewModel(_workerService, _shellViewModel);
        var messageTraceViewModel = new MessageTraceViewModel(_workerService, _shellViewModel);
        var mailFlowViewModel = new MailFlowViewModel(_workerService, _shellViewModel);


        sharedMailboxListViewModel.RecipientTypeFilter = "SharedMailbox";

                                                         
                                                                              
        _shellViewModel.Dashboard = dashboardViewModel;
        _shellViewModel.Mailboxes = mailboxListViewModel;
        _shellViewModel.DeletedMailboxes = deletedMailboxViewModel;
        _shellViewModel.SharedMailboxes = sharedMailboxListViewModel;
        _shellViewModel.MailboxDetails = mailboxDetailsViewModel;
        _shellViewModel.MailboxSpace = mailboxSpaceViewModel;
        _shellViewModel.DistributionLists = distributionListViewModel;
        _shellViewModel.Logs = logsViewModel;
        _shellViewModel.Tools = toolsViewModel;
        _shellViewModel.MessageTrace = messageTraceViewModel;
        _shellViewModel.MailFlow = mailFlowViewModel;
        _shellViewModel.RegisterNavigationStateSource(dashboardViewModel);
        _shellViewModel.RegisterNavigationStateSource(mailboxListViewModel);
        _shellViewModel.RegisterNavigationStateSource(deletedMailboxViewModel);
        _shellViewModel.RegisterNavigationStateSource(sharedMailboxListViewModel);
        _shellViewModel.RegisterNavigationStateSource(mailboxDetailsViewModel);
        _shellViewModel.RegisterNavigationStateSource(mailboxSpaceViewModel);
        _shellViewModel.RegisterNavigationStateSource(distributionListViewModel);

                                                                                     
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
                case NavigationPage.DeletedMailboxes:
                    await deletedMailboxViewModel.LoadAsync();
                    break;
                case NavigationPage.SharedMailboxes:
                    sharedMailboxListViewModel.RecipientTypeFilter = "SharedMailbox";
                    await sharedMailboxListViewModel.LoadAsync();
                    break;
                case NavigationPage.DistributionLists:
                    await distributionListViewModel.LoadAsync();
                    break;
                case NavigationPage.Tools:
                    await toolsViewModel.LoadAsync();
                    break;
                case NavigationPage.MessageTrace:
                    await messageTraceViewModel.LoadAsync();
                    break;
                case NavigationPage.MailFlow:
                    await mailFlowViewModel.LoadAsync();
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
                    case NavigationPage.DeletedMailboxes:
                        await deletedMailboxViewModel.LoadAsync();
                        break;
                    case NavigationPage.SharedMailboxes:
                        await sharedMailboxListViewModel.LoadAsync();
                        break;
                    case NavigationPage.DistributionLists:
                        await distributionListViewModel.LoadAsync();
                        break;
                    case NavigationPage.MessageTrace:
                        await messageTraceViewModel.LoadAsync();
                        break;
                    case NavigationPage.MailFlow:
                        await mailFlowViewModel.LoadAsync();
                        break;
                }
            }
        };

        MainWindow = mainWindow;
        mainWindow.Show();
        _ = _shellViewModel.StartWorkerOnStartupAsync();
    }

    protected override void OnExit(ExitEventArgs e)
    {
                                                                                

        try
        {
            _shellViewModel?.Dispose();
            _workerService?.StopWorkerAsync().GetAwaiter().GetResult();

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
