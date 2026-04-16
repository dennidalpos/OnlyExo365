using System.Reflection;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.ViewModels;

namespace ExchangeAdmin.Tests;

public class WorkerServiceFacetInterfaceTests
{
    [Theory]
    [InlineData(typeof(ShellViewModel), typeof(IConnectionWorkerService))]
    [InlineData(typeof(ShellConnectionStateViewModel), typeof(IConnectionWorkerService))]
    [InlineData(typeof(DashboardViewModel), typeof(IDashboardWorkerService))]
    [InlineData(typeof(ToolsViewModel), typeof(ISystemWorkerService))]
    [InlineData(typeof(ResourcesViewModel), typeof(IResourcesWorkerService))]
    [InlineData(typeof(ResourcesPageViewModel), typeof(IResourcesWorkerService))]
    [InlineData(typeof(ResourcesListStateViewModel), typeof(IResourcesWorkerService))]
    [InlineData(typeof(MailboxListViewModel), typeof(IMailboxesWorkerService))]
    [InlineData(typeof(MailboxProvisioningCandidatesViewModel), typeof(IMailboxesWorkerService))]
    [InlineData(typeof(MailboxDetailsViewModel), typeof(IMailboxesWorkerService))]
    [InlineData(typeof(MailboxSettingsEditorViewModel), typeof(IMailboxesWorkerService))]
    [InlineData(typeof(MailboxPermissionsEditorViewModel), typeof(IMailboxesWorkerService))]
    [InlineData(typeof(MailboxLicensesViewModel), typeof(IMailboxesWorkerService))]
    [InlineData(typeof(MailboxRestoreViewModel), typeof(IMailboxesWorkerService))]
    [InlineData(typeof(MailboxSpaceViewModel), typeof(IMailboxesWorkerService))]
    [InlineData(typeof(MailboxAccessReportViewModel), typeof(IMailboxesWorkerService))]
    [InlineData(typeof(DistributionListViewModel), typeof(IDistributionListsWorkerService))]
    [InlineData(typeof(DistributionListListViewModel), typeof(IDistributionListsWorkerService))]
    [InlineData(typeof(DistributionListDetailsViewModel), typeof(IDistributionListsWorkerService))]
    [InlineData(typeof(DistributionListSettingsEditorViewModel), typeof(IDistributionListsWorkerService))]
    [InlineData(typeof(DistributionListEditorViewModel), typeof(IDistributionListsWorkerService))]
    [InlineData(typeof(MigrationViewModel), typeof(IMigrationWorkerService))]
    [InlineData(typeof(ComplianceViewModel), typeof(IComplianceWorkerService))]
    [InlineData(typeof(MailSecurityViewModel), typeof(IMailSecurityWorkerService))]
    [InlineData(typeof(MailFlowViewModel), typeof(IMailFlowWorkerService))]
    public void Constructor_UsesExpectedWorkerFacet(Type viewModelType, Type expectedFacetType)
    {
        var constructor = viewModelType
            .GetConstructors(BindingFlags.Public | BindingFlags.Instance)
            .Single();

        var workerParameter = constructor.GetParameters().First();

        Assert.Equal(expectedFacetType, workerParameter.ParameterType);
    }

    [Fact]
    public void ConnectionWorkerFacet_ExposesWorkerConsoleVisibilityToggle()
    {
        var method = typeof(IConnectionWorkerService).GetMethod(nameof(IConnectionWorkerService.SetWorkerConsoleVisibilityAsync));

        Assert.NotNull(method);
        Assert.Equal(typeof(Task<Domain.Results.Result<SetWorkerConsoleVisibilityResponse>>), method!.ReturnType);
    }
}
