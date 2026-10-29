using System.Reflection;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;
using OnlyExo365.Shell.Ipc;
using OnlyExo365.Shell.Services;
using OnlyExo365.Shell.ViewModels;

namespace OnlyExo365.Tests;

public class SessionCacheRegressionTests
{
    [WpfTheory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SessionChange_RejectsLateResultsAndRereadsBothCaches(bool restartWorker)
    {
        var worker = new Worker();
        using var shell = new ShellViewModel(worker, new NavigationService());
        Set(shell, nameof(shell.WorkerState), WorkerConnectionState.Connected);
        Set(shell, nameof(shell.ExchangeState), ConnectionState.Connected);
        Set(shell, nameof(shell.ConnectedOrganization), "A");
        var cache = new CacheService();
        var dashboard = new DashboardViewModel(worker, new NavigationService(), shell, cache);
        var editor = new MailboxSettingsEditorViewModel(worker, shell, cache);
        var lateDashboard = dashboard.LoadAsync();
        var latePolicies = editor.LoadRetentionPoliciesAsync(default);
        if (restartWorker) Set(shell, nameof(shell.WorkerState), WorkerConnectionState.Stopped);
        else Set(shell, nameof(shell.ExchangeState), ConnectionState.Disconnected);
        Set(shell, nameof(shell.ConnectedOrganization), "B");
        Set(shell, nameof(shell.WorkerState), WorkerConnectionState.Connected);
        Set(shell, nameof(shell.ExchangeState), ConnectionState.Connected);
        worker.Dashboard.TrySetResult(Result<DashboardStatsDto>.Success(Stats(111)));
        worker.Policies.TrySetResult(Result<GetRetentionPoliciesResponse>.Success(new() { Policies = [new() { Name = "A" }] }));
        await Task.WhenAll(lateDashboard, latePolicies);
        Assert.Null(dashboard.Stats);
        Assert.Empty(editor.AvailableRetentionPolicies);
        worker.Dashboard = new();
        worker.Policies = new();
        worker.Dashboard.SetResult(Result<DashboardStatsDto>.Success(Stats(222)));
        worker.Policies.SetResult(Result<GetRetentionPoliciesResponse>.Success(new() { Policies = [new() { Name = "B" }] }));
        await dashboard.LoadAsync();
        await editor.LoadRetentionPoliciesAsync(default);
        Assert.Equal(222, dashboard.Stats!.MailboxCounts.UserMailboxes);
        Assert.Contains(editor.AvailableRetentionPolicies, policy => policy.Name == "B");
        await dashboard.LoadAsync();
        await editor.LoadRetentionPoliciesAsync(default);
        Assert.Equal(2, worker.DashboardCalls);
        Assert.Equal(2, worker.PolicyCalls);
        // Already-populated caches and visible data must also be cleared before the TTL expires.
        Set(shell, nameof(shell.ExchangeState), ConnectionState.Disconnected);
        Assert.Null(dashboard.Stats);
        Assert.Empty(editor.AvailableRetentionPolicies);
        Set(shell, nameof(shell.ConnectedOrganization), "C");
        Set(shell, nameof(shell.ExchangeState), ConnectionState.Connected);
        worker.Organization = "C";
        worker.Dashboard = new();
        worker.Policies = new();
        worker.Dashboard.SetResult(Result<DashboardStatsDto>.Success(Stats(333)));
        worker.Policies.SetResult(Result<GetRetentionPoliciesResponse>.Success(new() { Policies = [new() { Name = "C" }] }));
        await dashboard.LoadAsync();
        await editor.LoadRetentionPoliciesAsync(default);
        Assert.Equal(333, dashboard.Stats!.MailboxCounts.UserMailboxes);
        Assert.DoesNotContain(editor.AvailableRetentionPolicies, policy => policy.Name is "A" or "B");
        Assert.Equal(3, worker.DashboardCalls);
        Assert.Equal(3, worker.PolicyCalls);
    }

    private static DashboardStatsDto Stats(int count) => new() { MailboxCounts = new() { UserMailboxes = count }, Licenses = [new() { SkuId = "offline" }] };
    private static void Set(ShellViewModel shell, string property, object value) => typeof(ShellViewModel).GetProperty(property)!.SetValue(shell, value);
    private sealed class Worker : TestWorkerServiceBase
    {
        public TaskCompletionSource<Result<DashboardStatsDto>> Dashboard = new();
        public TaskCompletionSource<Result<GetRetentionPoliciesResponse>> Policies = new();
        public int DashboardCalls, PolicyCalls;
        public string Organization = "B";
        public override Task<Result<DashboardStatsDto>> GetDashboardStatsAsync(GetDashboardStatsRequest? request = null, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        { DashboardCalls++; return Dashboard.Task; }
        public override Task<Result<GetRetentionPoliciesResponse>> GetRetentionPoliciesAsync(GetRetentionPoliciesRequest? request = null, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        { PolicyCalls++; return Policies.Task; }
        public override Task<Result<ConnectionStatusDto>> GetConnectionStatusAsync(CancellationToken cancellationToken = default)
            => Task.FromResult(Result<ConnectionStatusDto>.Success(new() { State = OnlyExo365.Contracts.Dtos.ConnectionState.Connected, Organization = Organization }));
    }
}
