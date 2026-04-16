using System.ComponentModel;
using ExchangeAdmin.Presentation.Services;
using ExchangeAdmin.Presentation.ViewModels;

namespace ExchangeAdmin.Presentation.Bootstrap;

internal sealed class AppPageLoadCoordinator : IDisposable
{
    private readonly ShellViewModel _shellViewModel;
    private readonly NavigationService _navigationService;
    private readonly IReadOnlyDictionary<NavigationPage, Func<Task>> _pageLoaders;

    public AppPageLoadCoordinator(
        ShellViewModel shellViewModel,
        NavigationService navigationService,
        IReadOnlyDictionary<NavigationPage, Func<Task>> pageLoaders)
    {
        _shellViewModel = shellViewModel;
        _navigationService = navigationService;
        _pageLoaders = pageLoaders;

        _navigationService.PageChanged += OnPageChanged;
        _shellViewModel.PropertyChanged += OnShellPropertyChanged;
    }

    public void Dispose()
    {
        _navigationService.PageChanged -= OnPageChanged;
        _shellViewModel.PropertyChanged -= OnShellPropertyChanged;
    }

    private async void OnPageChanged(object? sender, NavigationPage page)
    {
        await LoadPageAsync(page);
    }

    private async void OnShellPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ShellViewModel.IsExchangeConnected) && _shellViewModel.IsExchangeConnected)
        {
            await LoadPageAsync(_navigationService.CurrentPage);
        }
    }

    private async Task LoadPageAsync(NavigationPage page)
    {
        try
        {
            if (_pageLoaders.TryGetValue(page, out var loader))
            {
                await loader();
            }
        }
        finally
        {
            _navigationService.CompleteNavigation(page);
        }
    }
}
