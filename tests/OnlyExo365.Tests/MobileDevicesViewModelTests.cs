using System.Reflection;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;
using OnlyExo365.Shell.Ipc;
using OnlyExo365.Shell.ViewModels;

namespace OnlyExo365.Tests;

public sealed class MobileDevicesViewModelTests
{
    [Fact]
    public async Task LoadAsync_PreloadsOnlyDeviceList()
    {
        var worker = new MobileDevicesTestWorkerService();
        worker.DeviceResponses.Enqueue(CreateDeviceResponse(1, totalCount: 1, hasMore: false));
        using var shell = new ShellViewModel(worker, new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new MobileDevicesViewModel(worker, shell);

        await viewModel.LoadAsync();

        Assert.Equal(1, worker.GetMobileDevicesCalls);
        Assert.Equal(0, worker.GetMobileDeviceDetailsCalls);
        Assert.Equal(0, worker.GetMobileDeviceMailboxPoliciesCalls);
        Assert.Single(viewModel.Devices);
    }

    [Fact]
    public async Task SelectingDevice_LoadsDetailsAndPoliciesOnDemand()
    {
        var worker = new MobileDevicesTestWorkerService();
        worker.DeviceResponses.Enqueue(CreateDeviceResponse(1, totalCount: 1, hasMore: false));
        using var shell = new ShellViewModel(worker, new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new MobileDevicesViewModel(worker, shell);
        await viewModel.LoadAsync();

        viewModel.SelectedDevice = viewModel.Devices[0];

        await Task.WhenAll(
            worker.MobileDeviceDetailsRequested.Task.WaitAsync(TimeSpan.FromSeconds(2)),
            worker.MobileDevicePoliciesRequested.Task.WaitAsync(TimeSpan.FromSeconds(2)));

        Assert.Equal("Default Mobile Policy", viewModel.SelectedDevice?.CurrentMailboxPolicy);
        Assert.Equal("user@contoso.com", viewModel.SelectedDevice?.UserPrincipalName);
        Assert.Equal("Mario Rossi", viewModel.SelectedDevice?.MailboxDisplayName);
        Assert.Equal("policy-01", viewModel.SelectedMailboxPolicyIdentity);
    }

    [Fact]
    public async Task LoadAsync_UsesPageSize250ForMobileDevicesPagination()
    {
        var worker = new MobileDevicesTestWorkerService();
        using var shell = new ShellViewModel(worker, new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new MobileDevicesViewModel(worker, shell);

        await viewModel.LoadAsync();

        Assert.Single(worker.MobileDeviceRequests);
        Assert.Equal(250, worker.MobileDeviceRequests[0].PageSize);
    }

    [Fact]
    public async Task RefreshCommand_PreservesLoadedDeviceDepth()
    {
        var worker = new MobileDevicesTestWorkerService();
        worker.DeviceResponses.Enqueue(CreateDeviceResponse(250, totalCount: 540, hasMore: true));
        worker.DeviceResponses.Enqueue(CreateDeviceResponse(250, skip: 250, totalCount: 540, hasMore: true));
        worker.DeviceResponses.Enqueue(CreateDeviceResponse(500, totalCount: 540, hasMore: true));
        using var shell = new ShellViewModel(worker, new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new MobileDevicesViewModel(worker, shell);

        await viewModel.LoadAsync();
        viewModel.LoadMoreCommand.Execute(null);
        await worker.SecondMobileDeviceRequestReceived.Task.WaitAsync(TimeSpan.FromSeconds(2));
        viewModel.RefreshCommand.Execute(null);
        await worker.ThirdMobileDeviceRequestReceived.Task.WaitAsync(TimeSpan.FromSeconds(2));

        Assert.Equal([250, 250, 500], worker.MobileDeviceRequests.Select(request => request.PageSize));
        Assert.Equal(500, viewModel.Devices.Count);
    }

    private static void SetExchangeConnected(ShellViewModel shell)
    {
        var property = typeof(ShellViewModel)
            .GetProperty(nameof(ShellViewModel.ExchangeState), BindingFlags.Instance | BindingFlags.Public);

        Assert.NotNull(property);
        property!.SetValue(shell, ConnectionState.Connected);
    }

    private sealed class MobileDevicesTestWorkerService : TestWorkerServiceBase
    {
        public int GetMobileDevicesCalls { get; private set; }
        public int GetMobileDeviceDetailsCalls { get; private set; }
        public int GetMobileDeviceMailboxPoliciesCalls { get; private set; }
        public List<GetMobileDevicesRequest> MobileDeviceRequests { get; } = [];
        public Queue<GetMobileDevicesResponse> DeviceResponses { get; } = new();
        public TaskCompletionSource MobileDeviceDetailsRequested { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public TaskCompletionSource MobileDevicePoliciesRequested { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public TaskCompletionSource SecondMobileDeviceRequestReceived { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public TaskCompletionSource ThirdMobileDeviceRequestReceived { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public override Task<Result<CapabilityMapDto>> DetectCapabilitiesAsync(bool forceRefresh = false, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            return Task.FromResult(Result<CapabilityMapDto>.Success(new CapabilityMapDto
            {
                Features = new FeatureCapabilitiesDto
                {
                    CanGetMobileDevice = true,
                    CanGetMobileDeviceStatistics = true,
                    CanGetCasMailbox = true,
                    CanGetMobileDeviceMailboxPolicy = true,
                    CanSetCasMailbox = true
                }
            }));
        }

        public override Task<Result<GetMobileDevicesResponse>> GetMobileDevicesAsync(GetMobileDevicesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            GetMobileDevicesCalls++;
            MobileDeviceRequests.Add(new GetMobileDevicesRequest
            {
                SearchQuery = request.SearchQuery,
                AccessState = request.AccessState,
                PageSize = request.PageSize,
                Skip = request.Skip,
                SortBy = request.SortBy,
                SortDescending = request.SortDescending
            });
            if (MobileDeviceRequests.Count == 2)
            {
                SecondMobileDeviceRequestReceived.TrySetResult();
            }
            else if (MobileDeviceRequests.Count == 3)
            {
                ThirdMobileDeviceRequestReceived.TrySetResult();
            }
            var response = DeviceResponses.Count > 0
                ? DeviceResponses.Dequeue()
                : CreateDeviceResponse(request.PageSize, request.Skip, totalCount: 1, hasMore: false);

            response.PageSize = request.PageSize;
            response.Skip = request.Skip;
            return Task.FromResult(Result<GetMobileDevicesResponse>.Success(response));
        }

        public override Task<Result<GetMobileDeviceDetailsResponse>> GetMobileDeviceDetailsAsync(GetMobileDeviceDetailsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            GetMobileDeviceDetailsCalls++;
            MobileDeviceDetailsRequested.TrySetResult();
            eventHandler?.Invoke(new EventEnvelope
            {
                CorrelationId = "mobile-device-detail",
                EventType = EventType.Progress,
                Payload = JsonMessageSerializer.ToJsonElement(new ProgressEventPayload
                {
                    PercentComplete = 50,
                    StatusMessage = "Mobile device statistics loaded",
                    CurrentItem = 1,
                    TotalItems = 2
                })
            });

            return Task.FromResult(Result<GetMobileDeviceDetailsResponse>.Success(new GetMobileDeviceDetailsResponse
            {
                Device = new MobileDeviceListItemDto
                {
                    Identity = request.Identity,
                    DeviceId = "dev-01",
                    UserDisplayName = "Mario Rossi",
                    UserPrincipalName = "user@contoso.com",
                    MailboxIdentity = "user@contoso.com",
                    MailboxDisplayName = "Mario Rossi",
                    DeviceAccessState = "Allowed",
                    DeviceType = "Phone",
                    CurrentMailboxPolicy = "Default Mobile Policy",
                    LastSuccessSync = new DateTime(2026, 3, 12, 8, 30, 0, DateTimeKind.Utc)
                }
            }));
        }

        public override Task<Result<GetMobileDeviceMailboxPoliciesResponse>> GetMobileDeviceMailboxPoliciesAsync(GetMobileDeviceMailboxPoliciesRequest? request = null, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            GetMobileDeviceMailboxPoliciesCalls++;
            MobileDevicePoliciesRequested.TrySetResult();
            return Task.FromResult(Result<GetMobileDeviceMailboxPoliciesResponse>.Success(new GetMobileDeviceMailboxPoliciesResponse
            {
                Policies =
                [
                    new MobileDeviceMailboxPolicyDto
                    {
                        Identity = "policy-01",
                        Name = "Default Mobile Policy"
                    }
                ]
            }));
        }
    }

    private static GetMobileDevicesResponse CreateDeviceResponse(int pageSize, int skip = 0, int totalCount = 1, bool hasMore = false)
    {
        var devices = Enumerable.Range(skip, pageSize)
            .Select(index => new MobileDeviceListItemDto
            {
                Identity = $"device-{index:D4}",
                DeviceId = $"dev-{index:D4}",
                UserDisplayName = $"User {index:D4}",
                MailboxIdentity = $"mailbox-{index:D4}",
                MailboxDisplayName = $"Mailbox {index:D4}",
                DeviceAccessState = "Allowed",
                DeviceType = "Phone"
            })
            .ToList();

        return new GetMobileDevicesResponse
        {
            Devices = devices,
            TotalCount = totalCount,
            PageSize = pageSize,
            Skip = skip,
            HasMore = hasMore,
            IsTotalCountExact = true
        };
    }
}

