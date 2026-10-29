using OnlyExo365.Shell.Services;
using OnlyExo365.Shell.ViewModels;

namespace OnlyExo365.Tests;

public sealed class NavigationServiceTests
{
    [Fact]
    public void ShellNavigationState_RaisesEveryPublicPageBindingWhenPageChanges()
    {
        var navigation = new NavigationService();
        using var state = new ShellNavigationStateViewModel(
            navigation,
            new ShellProgressViewModel(),
            new ShellPromptViewModel());
        var changedProperties = new List<string>();
        var expectedPageBindings = typeof(ShellNavigationStateViewModel)
            .GetProperties()
            .Where(property => property.PropertyType == typeof(bool))
            .Select(property => property.Name)
            .Where(name => name.StartsWith("Is", StringComparison.Ordinal) &&
                           name.EndsWith("Page", StringComparison.Ordinal))
            .ToArray();

        state.PropertyChanged += (_, args) =>
        {
            if (!string.IsNullOrWhiteSpace(args.PropertyName))
            {
                changedProperties.Add(args.PropertyName);
            }
        };

        navigation.NavigateTo(NavigationPage.Contacts);

        Assert.Equal(NavigationPage.Contacts, state.CurrentPage);
        Assert.False(state.IsDashboardPage);
        Assert.True(state.IsContactsPage);
        foreach (var binding in expectedPageBindings)
        {
            Assert.Contains(binding, changedProperties);
        }

        Assert.Contains(nameof(ShellNavigationStateViewModel.CurrentPageTitle), changedProperties);
    }

    [Fact]
    public void NavigateTo_WhenCancelled_DoesNotChangePageOrSelection()
    {
        var navigation = new NavigationService();
        navigation.NavigateToDetails(NavigationPage.Mailboxes, "mailbox-1", new object());
        navigation.CompleteNavigation(NavigationPage.Mailboxes);
        navigation.Navigating += (_, args) => args.Cancel = true;

        navigation.NavigateTo(NavigationPage.Contacts);

        Assert.Equal(NavigationPage.Mailboxes, navigation.CurrentPage);
        Assert.Equal("mailbox-1", navigation.SelectedIdentity);
        Assert.NotNull(navigation.SelectedItem);
        Assert.False(navigation.IsNavigationPending);
    }

    [Fact]
    public void NavigateToDetails_WhenCancelled_DoesNotChangePageOrSelection()
    {
        var navigation = new NavigationService();
        navigation.NavigateToDetails(NavigationPage.Mailboxes, "mailbox-1", new object());
        navigation.CompleteNavigation(NavigationPage.Mailboxes);
        navigation.Navigating += (_, args) => args.Cancel = true;

        navigation.NavigateToDetails(NavigationPage.DistributionLists, "group-1", new object());

        Assert.Equal(NavigationPage.Mailboxes, navigation.CurrentPage);
        Assert.Equal("mailbox-1", navigation.SelectedIdentity);
        Assert.NotNull(navigation.SelectedItem);
        Assert.False(navigation.IsNavigationPending);
    }
}
