using System.Reflection;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;
using ExchangeAdmin.Infrastructure.Ipc;
using ExchangeAdmin.Presentation.Services;
using ExchangeAdmin.Presentation.ViewModels;

namespace ExchangeAdmin.Tests;

public sealed class NavigationLoadingCoordinatorTests
{
    [Fact]
    public void ShellViewModel_LocksNavigationImmediatelyWhenSectionChangeStarts()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());

        shell.NavigationService.NavigateTo(NavigationPage.Contacts);

        Assert.Equal(NavigationPage.Contacts, shell.CurrentPage);
        Assert.True(shell.IsNavigationLocked);
        Assert.False(shell.CanNavigate);
        Assert.True(shell.NavigationService.IsNavigationPending);
    }

    [Fact]
    public async Task AppPageLoadCoordinator_ReleasesPendingNavigationWhenLoaderCompletes()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        var loaderStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var loaderGate = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        using var coordinator = CreateCoordinator(
            shell,
            shell.NavigationService,
            new Dictionary<NavigationPage, Func<Task>>
            {
                [NavigationPage.Contacts] = async () =>
                {
                    loaderStarted.SetResult();
                    await loaderGate.Task;
                }
            });

        shell.NavigationService.NavigateTo(NavigationPage.Contacts);
        await loaderStarted.Task.WaitAsync(TimeSpan.FromSeconds(2));

        Assert.Equal(NavigationPage.Contacts, shell.CurrentPage);
        Assert.True(shell.NavigationService.IsNavigationPending);
        Assert.True(shell.IsNavigationLocked);

        shell.NavigationService.NavigateTo(NavigationPage.Resources);

        Assert.Equal(NavigationPage.Contacts, shell.CurrentPage);
        Assert.Equal(NavigationPage.Contacts, shell.NavigationService.PendingPage);

        loaderGate.SetResult();
        await AssertPendingStateAsync(shell.NavigationService, expectedPending: false);

        Assert.False(shell.IsNavigationLocked);
        Assert.True(shell.CanNavigate);
        Assert.Null(shell.NavigationService.PendingPage);
    }

    private static async Task AssertPendingStateAsync(NavigationService navigationService, bool expectedPending)
    {
        var timeoutAt = DateTime.UtcNow.AddSeconds(2);
        while (navigationService.IsNavigationPending != expectedPending && DateTime.UtcNow < timeoutAt)
        {
            await Task.Delay(20);
        }

        Assert.Equal(expectedPending, navigationService.IsNavigationPending);
    }

    private static IDisposable CreateCoordinator(
        ShellViewModel shellViewModel,
        NavigationService navigationService,
        IReadOnlyDictionary<NavigationPage, Func<Task>> pageLoaders)
    {
        var coordinatorType = typeof(ShellViewModel).Assembly
            .GetType("ExchangeAdmin.Presentation.Bootstrap.AppPageLoadCoordinator");

        Assert.NotNull(coordinatorType);

        var instance = Activator.CreateInstance(
            coordinatorType!,
            BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public,
            binder: null,
            args: [shellViewModel, navigationService, pageLoaders],
            culture: null);

        Assert.NotNull(instance);
        return (IDisposable)instance!;
    }

}
