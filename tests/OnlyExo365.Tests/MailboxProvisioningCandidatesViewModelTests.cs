using System.Reflection;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;
using OnlyExo365.Shell.Ipc;
using OnlyExo365.Shell.ViewModels;

namespace OnlyExo365.Tests;

public sealed class MailboxProvisioningCandidatesViewModelTests
{
    [Fact]
    public async Task LoadAsync_UsesDefaultFiltersAndLoadsAvailableLicenses()
    {
        var worker = new MailboxProvisioningWorkerService();
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new MailboxProvisioningCandidatesViewModel(worker, shell);

        await viewModel.LoadAsync();

        Assert.Single(worker.ProvisioningRequests);
        Assert.True(worker.ProvisioningRequests[0].OnlyWithoutLicense);
        Assert.True(worker.ProvisioningRequests[0].OnlyWithoutMail);
        Assert.Equal(250, worker.ProvisioningRequests[0].PageSize);
        Assert.Single(viewModel.Candidates);
        Assert.Equal(1, worker.GetAvailableLicensesCalls);
    }

    [Fact]
    public async Task SelectingCandidate_LoadsAssignedLicensesForUserPrincipalName()
    {
        var worker = new MailboxProvisioningWorkerService();
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new MailboxProvisioningCandidatesViewModel(worker, shell);

        await viewModel.LoadAsync();
        await WaitForConditionAsync(() => worker.UserLicenseRequests.Count == 1);

        Assert.Equal("mario.rossi@contoso.com", worker.UserLicenseRequests[0].UserPrincipalName);
        Assert.Empty(viewModel.Licenses.AssignedLicenses);
    }

    [Fact]
    public async Task AddLicenseCommand_RefreshesProvisioningListAfterMutation()
    {
        var worker = new MailboxProvisioningWorkerService();
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        ErrorDialogService.ConfirmationHandlerOverride = (_, _) => true;
        var viewModel = new MailboxProvisioningCandidatesViewModel(worker, shell);

        try
        {
            await viewModel.LoadAsync();
            await WaitForConditionAsync(() => worker.UserLicenseRequests.Count == 1);

            viewModel.Licenses.SelectedLicenseToAdd = viewModel.Licenses.AvailableLicenses[0];
            viewModel.Licenses.AddLicenseCommand.Execute(null);

            await WaitForConditionAsync(() =>
                worker.SetUserLicenseCalls == 1 &&
                worker.ProvisioningRequests.Count >= 2 &&
                viewModel.Candidates[0].HasAssignedLicense);

            Assert.Equal("mario.rossi@contoso.com", worker.LastSetUserLicenseRequest?.UserPrincipalName);
        }
        finally
        {
            ErrorDialogService.ConfirmationHandlerOverride = null;
        }
    }

    private static async Task WaitForConditionAsync(Func<bool> condition)
    {
        for (var attempt = 0; attempt < 40; attempt++)
        {
            if (condition())
            {
                return;
            }

            await Task.Delay(25);
        }

        Assert.True(condition(), "Condition not reached in time.");
    }

    private static void SetExchangeConnected(ShellViewModel shell)
    {
        var property = typeof(ShellViewModel)
            .GetProperty(nameof(ShellViewModel.ExchangeState), BindingFlags.Instance | BindingFlags.Public);

        Assert.NotNull(property);
        property!.SetValue(shell, ConnectionState.Connected);
    }

    private sealed class MailboxProvisioningWorkerService : TestMailboxesWorkerServiceBase
    {
        public List<GetMailboxProvisioningCandidatesRequest> ProvisioningRequests { get; } = [];
        public List<GetUserLicensesRequest> UserLicenseRequests { get; } = [];
        public int GetAvailableLicensesCalls { get; private set; }
        public int SetUserLicenseCalls { get; private set; }
        public SetUserLicenseRequest? LastSetUserLicenseRequest { get; private set; }

        public override Task<Result<GetMailboxProvisioningCandidatesResponse>> GetMailboxProvisioningCandidatesAsync(GetMailboxProvisioningCandidatesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            ProvisioningRequests.Add(new GetMailboxProvisioningCandidatesRequest
            {
                SearchQuery = request.SearchQuery,
                OnlyWithoutLicense = request.OnlyWithoutLicense,
                OnlyWithoutMail = request.OnlyWithoutMail,
                PageSize = request.PageSize,
                Skip = request.Skip
            });

            var hasLicense = SetUserLicenseCalls > 0;
            return Task.FromResult(Result<GetMailboxProvisioningCandidatesResponse>.Success(new GetMailboxProvisioningCandidatesResponse
            {
                Candidates =
                [
                    new MailboxProvisioningCandidateDto
                    {
                        DisplayName = "Mario Rossi",
                        UserPrincipalName = "mario.rossi@contoso.com",
                        Mail = null,
                        AccountEnabled = false,
                        HasAssignedLicense = hasLicense,
                        HasMailAddress = false
                    }
                ],
                TotalCount = 1,
                Skip = request.Skip,
                PageSize = request.PageSize,
                HasMore = false
            }));
        }

        public override Task<Result<GetUserLicensesResponse>> GetUserLicensesAsync(GetUserLicensesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            UserLicenseRequests.Add(new GetUserLicensesRequest
            {
                UserPrincipalName = request.UserPrincipalName
            });

            List<UserLicenseDto> licenses = SetUserLicenseCalls > 0
                ? [new UserLicenseDto { SkuId = "sku-01", SkuPartNumber = "ENTERPRISEPACK", DisplayName = "Office 365 E3" }]
                : [];

            return Task.FromResult(Result<GetUserLicensesResponse>.Success(new GetUserLicensesResponse
            {
                Licenses = licenses
            }));
        }

        public override Task<Result<GetAvailableLicensesResponse>> GetAvailableLicensesAsync(Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            GetAvailableLicensesCalls++;
            return Task.FromResult(Result<GetAvailableLicensesResponse>.Success(new GetAvailableLicensesResponse
            {
                Licenses =
                [
                    new TenantLicenseDto
                    {
                        SkuId = "sku-01",
                        SkuPartNumber = "ENTERPRISEPACK",
                        DisplayName = "Office 365 E3",
                        Available = 10
                    }
                ]
            }));
        }

        public override Task<Result> SetUserLicenseAsync(SetUserLicenseRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            SetUserLicenseCalls++;
            LastSetUserLicenseRequest = request;
            return Task.FromResult(Result.Success());
        }
    }
}

