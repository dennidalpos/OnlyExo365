using System.Reflection;
using OnlyExo365.Shell.Security;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts;
using OnlyExo365.Shell.ViewModels;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Errors;
using OnlyExo365.Contracts.Results;
using OnlyExo365.Shell.Ipc;

namespace OnlyExo365.Tests;

public sealed class MailboxLicensesViewModelTests
{
    [Fact]
    public void DefaultConfiguration_BlocksLicenseWriteActionsWhenLeastPrivilegeEvaluationIsBlocked()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        SetCapabilities(shell, new CapabilityMapDto());

        var viewModel = new MailboxLicensesViewModel(
            new MailboxLicensesWorkerService(),
            shell,
            getPrimarySmtpAddress: () => "user@contoso.com",
            getDisplayName: () => "User Contoso");

        viewModel.SelectedLicenseToAdd = new TenantLicenseDto
        {
            SkuId = "sku-01",
            SkuPartNumber = "ENTERPRISEPACK"
        };

        Assert.True(viewModel.IsLicenseWriteBlocked);
        Assert.Contains("LicenseAssignment.ReadWrite.All", viewModel.LicenseWriteDisabledMessage, StringComparison.Ordinal);
        Assert.False(viewModel.AddLicenseCommand.CanExecute(null));
        Assert.False(viewModel.RemoveLicenseCommand.CanExecute(new UserLicenseDto { SkuId = "sku-01", SkuPartNumber = "ENTERPRISEPACK" }));
    }

    [Fact]
    public void OptInConfiguration_AllowsLicenseWriteActionsWhenWriteScopeIsConfigured()
    {
        using var shell = new ShellViewModel(
            new ConnectedConnectionWorkerServiceStub(),
            new NavigationService(),
            new ExchangeOnlineConfiguration
            {
                GraphLicenseWriteScopes =
                [
                    "LicenseAssignment.ReadWrite.All"
                ]
            });
        SetExchangeConnected(shell);
        SetCapabilities(shell, new CapabilityMapDto());

        var viewModel = new MailboxLicensesViewModel(
            new MailboxLicensesWorkerService(),
            shell,
            getPrimarySmtpAddress: () => "user@contoso.com",
            getDisplayName: () => "User Contoso");

        viewModel.SelectedLicenseToAdd = new TenantLicenseDto
        {
            SkuId = "sku-01",
            SkuPartNumber = "ENTERPRISEPACK"
        };

        Assert.False(viewModel.IsLicenseWriteBlocked);
        Assert.True(viewModel.AddLicenseCommand.CanExecute(null));
        Assert.True(viewModel.RemoveLicenseCommand.CanExecute(new UserLicenseDto { SkuId = "sku-01", SkuPartNumber = "ENTERPRISEPACK" }));
    }

    [Fact]
    public void BuildUsageLocationSuggestionMessage_ShowsSourceAndSuggestedValue()
    {
        var message = MailboxLicensesViewModel.BuildUsageLocationSuggestionMessage(new GetUsageLocationSuggestionResponse
        {
            SuggestedUsageLocation = "DE",
            SuggestionSource = "Tenant",
            SuggestionDetails = "Suggested from Microsoft Graph organization.countryLetterCode."
        });

        Assert.Contains("DE", message, StringComparison.Ordinal);
        Assert.Contains("tenant", message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void UsageLocations_ExposeCompleteTwoLetterCountryCodeList()
    {
        using var shell = new ShellViewModel(
            new ConnectedConnectionWorkerServiceStub(),
            new NavigationService(),
            new ExchangeOnlineConfiguration
            {
                GraphLicenseWriteScopes =
                [
                    "LicenseAssignment.ReadWrite.All"
                ]
            });
        SetExchangeConnected(shell);
        SetCapabilities(shell, new CapabilityMapDto());

        var viewModel = new MailboxLicensesViewModel(
            new MailboxLicensesWorkerService(),
            shell,
            getPrimarySmtpAddress: () => "user@contoso.com",
            getDisplayName: () => "User Contoso");

        Assert.True(viewModel.UsageLocations.Count > 200);
        Assert.Contains(viewModel.UsageLocations, option => option.Code == "DE");
        Assert.Contains(viewModel.UsageLocations, option => option.Code == "IT");
        Assert.Contains(viewModel.UsageLocations, option => option.Code == "US");
    }

    [Fact]
    public async Task AddLicenseAsync_LoadsUsageLocationSuggestionFromTenantWhenWorkerReturnsUsageLocationError()
    {
        var worker = new MailboxLicensesUsageLocationErrorWorkerService();
        using var shell = new ShellViewModel(
            new ConnectedConnectionWorkerServiceStub(),
            new NavigationService(),
            new ExchangeOnlineConfiguration
            {
                GraphLicenseWriteScopes =
                [
                    "LicenseAssignment.ReadWrite.All"
                ]
            });
        SetExchangeConnected(shell);
        SetCapabilities(shell, new CapabilityMapDto());

        var viewModel = new MailboxLicensesViewModel(
            worker,
            shell,
            getPrimarySmtpAddress: () => "user@contoso.com",
            getDisplayName: () => "User Contoso");

        viewModel.SelectedLicenseToAdd = new TenantLicenseDto
        {
            SkuId = "sku-01",
            SkuPartNumber = "ENTERPRISEPACK"
        };

        viewModel.AddLicenseCommand.Execute(null);

        await WaitForConditionAsync(() => !string.IsNullOrWhiteSpace(viewModel.UsageLocationSuggestionMessage));

        Assert.NotNull(viewModel.LicenseErrorMessage);
        Assert.Contains("DE", viewModel.UsageLocationSuggestionMessage, StringComparison.Ordinal);
        Assert.Contains("tenant", viewModel.UsageLocationSuggestionMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("DE", viewModel.SelectedUsageLocation?.Code);
        Assert.Equal(1, worker.GetUsageLocationSuggestionCalls);
    }

    [Fact]
    public async Task ApplySuggestedUsageLocationCommand_UpdatesUsageLocationAndRetriesLicenseAssignment()
    {
        var worker = new MailboxLicensesUsageLocationErrorWorkerService();
        using var shell = new ShellViewModel(
            new ConnectedConnectionWorkerServiceStub(),
            new NavigationService(),
            new ExchangeOnlineConfiguration
            {
                GraphLicenseWriteScopes =
                [
                    "LicenseAssignment.ReadWrite.All"
                ]
            });
        SetExchangeConnected(shell);
        SetCapabilities(shell, new CapabilityMapDto());

        var viewModel = new MailboxLicensesViewModel(
            worker,
            shell,
            getPrimarySmtpAddress: () => "user@contoso.com",
            getDisplayName: () => "User Contoso");

        viewModel.SelectedLicenseToAdd = new TenantLicenseDto
        {
            SkuId = "sku-01",
            SkuPartNumber = "ENTERPRISEPACK"
        };

        viewModel.AddLicenseCommand.Execute(null);
        await WaitForConditionAsync(() => !string.IsNullOrWhiteSpace(viewModel.UsageLocationSuggestionMessage));

        worker.AllowLicenseAssignment = true;
        viewModel.ApplySuggestedUsageLocationCommand.Execute(null);

        await WaitForConditionAsync(() => worker.SetUserUsageLocationCalls == 1 && worker.SetUserLicenseCalls >= 2);

        Assert.Equal("DE", worker.LastSetUserUsageLocationRequest?.UsageLocation);
        Assert.Null(viewModel.UsageLocationSuggestionMessage);
        Assert.Null(viewModel.LicenseErrorMessage);
    }

    [Fact]
    public async Task ApplySuggestedUsageLocationCommand_UsesManualSelectionWithoutPendingSuggestion()
    {
        var worker = new MailboxLicensesUsageLocationUpdateWorkerService();
        using var shell = new ShellViewModel(
            new ConnectedConnectionWorkerServiceStub(),
            new NavigationService(),
            new ExchangeOnlineConfiguration
            {
                GraphLicenseWriteScopes =
                [
                    "LicenseAssignment.ReadWrite.All"
                ]
            });
        SetExchangeConnected(shell);
        SetCapabilities(shell, new CapabilityMapDto());

        var viewModel = new MailboxLicensesViewModel(
            worker,
            shell,
            getPrimarySmtpAddress: () => "user@contoso.com",
            getDisplayName: () => "User Contoso");

        viewModel.SelectedUsageLocation = Assert.Single(viewModel.UsageLocations, option => option.Code == "IT");
        viewModel.ApplySuggestedUsageLocationCommand.Execute(null);

        await WaitForConditionAsync(() => worker.SetUserUsageLocationCalls == 1);

        Assert.Equal("IT", worker.LastSetUserUsageLocationRequest?.UsageLocation);
        Assert.Null(viewModel.LicenseErrorMessage);
    }

    [Fact]
    public async Task RemoveLicenseCommand_RemovesLicenseWithoutConfirmationPrompt()
    {
        var worker = new MailboxLicensesMutationWorkerService();
        using var shell = new ShellViewModel(
            new ConnectedConnectionWorkerServiceStub(),
            new NavigationService(),
            new ExchangeOnlineConfiguration
            {
                GraphLicenseWriteScopes =
                [
                    "LicenseAssignment.ReadWrite.All"
                ]
            });
        SetExchangeConnected(shell);
        SetCapabilities(shell, new CapabilityMapDto());

        var viewModel = new MailboxLicensesViewModel(
            worker,
            shell,
            getPrimarySmtpAddress: () => "user@contoso.com",
            getDisplayName: () => "User Contoso");

        var assignedLicense = new UserLicenseDto
        {
            SkuId = "sku-01",
            SkuPartNumber = "ENTERPRISEPACK"
        };

        viewModel.RemoveLicenseCommand.Execute(assignedLicense);

        await WaitForConditionAsync(() => worker.SetUserLicenseCalls == 1);

        Assert.Equal("user@contoso.com", worker.LastSetUserLicenseRequest?.UserPrincipalName);
        Assert.Equal("sku-01", Assert.Single(worker.LastSetUserLicenseRequest?.RemoveLicenseSkuIds ?? []));
        Assert.Null(viewModel.LicenseErrorMessage);
    }

    private static async Task WaitForConditionAsync(Func<bool> condition)
    {
        for (var attempt = 0; attempt < 80; attempt++)
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

    private static void SetCapabilities(ShellViewModel shell, CapabilityMapDto capabilities)
    {
        var property = typeof(ShellViewModel)
            .GetProperty(nameof(ShellViewModel.Capabilities), BindingFlags.Instance | BindingFlags.Public);

        Assert.NotNull(property);
        property!.SetValue(shell, capabilities);
    }

    private sealed class MailboxLicensesWorkerService : TestMailboxesWorkerServiceBase
    {
    }

    private sealed class MailboxLicensesUsageLocationErrorWorkerService : TestMailboxesWorkerServiceBase
    {
        public bool AllowLicenseAssignment { get; set; }
        public int GetUsageLocationSuggestionCalls { get; private set; }
        public int SetUserLicenseCalls { get; private set; }
        public int SetUserUsageLocationCalls { get; private set; }
        public SetUserUsageLocationRequest? LastSetUserUsageLocationRequest { get; private set; }

        public override Task<Result> SetUserLicenseAsync(SetUserLicenseRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            SetUserLicenseCalls++;
            if (AllowLicenseAssignment)
            {
                return Task.FromResult(Result.Success());
            }

            return Task.FromResult(Result.Failure(NormalizedError.Create(
                ErrorCode.InvalidParameter,
                "License assignment cannot be completed because the user usage location is not set. Set UsageLocation to a valid two-letter country/region code and retry.")));
        }

        public override Task<Result<GetUsageLocationSuggestionResponse>> GetUsageLocationSuggestionAsync(GetUsageLocationSuggestionRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            GetUsageLocationSuggestionCalls++;
            return Task.FromResult(Result<GetUsageLocationSuggestionResponse>.Success(new GetUsageLocationSuggestionResponse
            {
                UserPrincipalName = request.UserPrincipalName,
                SuggestedUsageLocation = "DE",
                SuggestionSource = "Tenant",
                SuggestionDetails = "Suggested from Microsoft Graph organization.countryLetterCode."
            }));
        }

        public override Task<Result> SetUserUsageLocationAsync(SetUserUsageLocationRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            SetUserUsageLocationCalls++;
            LastSetUserUsageLocationRequest = request;
            return Task.FromResult(Result.Success());
        }

        public override Task<Result<GetUserLicensesResponse>> GetUserLicensesAsync(GetUserLicensesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result<GetUserLicensesResponse>.Success(new GetUserLicensesResponse()));

        public override Task<Result<GetAvailableLicensesResponse>> GetAvailableLicensesAsync(Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result<GetAvailableLicensesResponse>.Success(new GetAvailableLicensesResponse()));
    }

    private sealed class MailboxLicensesMutationWorkerService : TestMailboxesWorkerServiceBase
    {
        public int SetUserLicenseCalls { get; private set; }
        public SetUserLicenseRequest? LastSetUserLicenseRequest { get; private set; }

        public override Task<Result> SetUserLicenseAsync(SetUserLicenseRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            SetUserLicenseCalls++;
            LastSetUserLicenseRequest = request;
            return Task.FromResult(Result.Success());
        }

        public override Task<Result<GetUserLicensesResponse>> GetUserLicensesAsync(GetUserLicensesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result<GetUserLicensesResponse>.Success(new GetUserLicensesResponse()));

        public override Task<Result<GetAvailableLicensesResponse>> GetAvailableLicensesAsync(Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result<GetAvailableLicensesResponse>.Success(new GetAvailableLicensesResponse()));
    }

    private sealed class MailboxLicensesUsageLocationUpdateWorkerService : TestMailboxesWorkerServiceBase
    {
        public int SetUserUsageLocationCalls { get; private set; }
        public SetUserUsageLocationRequest? LastSetUserUsageLocationRequest { get; private set; }

        public override Task<Result> SetUserUsageLocationAsync(SetUserUsageLocationRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            SetUserUsageLocationCalls++;
            LastSetUserUsageLocationRequest = request;
            return Task.FromResult(Result.Success());
        }

        public override Task<Result<GetUserLicensesResponse>> GetUserLicensesAsync(GetUserLicensesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result<GetUserLicensesResponse>.Success(new GetUserLicensesResponse()));

        public override Task<Result<GetAvailableLicensesResponse>> GetAvailableLicensesAsync(Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result<GetAvailableLicensesResponse>.Success(new GetAvailableLicensesResponse()));
    }
}

