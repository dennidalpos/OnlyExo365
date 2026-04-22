using System.Runtime.InteropServices;
using System.Windows;
using OnlyExo365.Contracts;
using OnlyExo365.Contracts.Diagnostics;
using OnlyExo365.Shell.Bootstrap;
using OnlyExo365.Shell.Localization;
using OnlyExo365.Shell.Services;

namespace OnlyExo365.Shell;

public partial class App : System.Windows.Application
{
    private AppRuntimeContext? _runtimeContext;
    private GlobalExceptionMonitor? _globalExceptionMonitor;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        _globalExceptionMonitor = new GlobalExceptionMonitor(new PersistentLogWriter("ui"));
        _globalExceptionMonitor.Register(this);

        ExchangeOnlineConfiguration exchangeConfiguration;
        try
        {
            exchangeConfiguration = ExchangeConfigurationLoader.Load();
        }
        catch (ExchangeConfigurationLoader.ExchangeConfigurationLoadException ex)
        {
            MessageBox.Show(
                ex.Message,
                Loc.Get("Msg.ConfigurationError"),
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            Shutdown(-1);
            return;
        }

        _runtimeContext = AppCompositionRoot.Create(exchangeConfiguration);

        MainWindow = _runtimeContext.MainWindow;
        MainWindow.Show();
        _ = _runtimeContext.StartAsync();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        try
        {
            _globalExceptionMonitor?.Dispose();

            if (_runtimeContext is not null)
            {
                _runtimeContext.DisposeAsync().AsTask().GetAwaiter().GetResult();
            }
        }
        catch
        {
        }
        finally
        {
            _globalExceptionMonitor?.Dispose();
            _globalExceptionMonitor = null;
            _runtimeContext = null;
        }

        base.OnExit(e);
    }
}

