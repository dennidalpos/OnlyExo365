using System.Runtime.InteropServices;
using System.Windows;
using ExchangeAdmin.Contracts;
using ExchangeAdmin.Contracts.Diagnostics;
using ExchangeAdmin.Presentation.Bootstrap;
using ExchangeAdmin.Presentation.Localization;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation;

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
