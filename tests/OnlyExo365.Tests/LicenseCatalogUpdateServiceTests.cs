using System.Net;
using System.Net.Http;
using OnlyExo365.Shell.Services;

namespace OnlyExo365.Tests;

public class LicenseCatalogUpdateServiceTests
{
    [Theory]
    [InlineData(CatalogAutoUpdateMode.Daily)]
    [InlineData(CatalogAutoUpdateMode.Monthly)]
    [InlineData(CatalogAutoUpdateMode.Disabled)]
    public async Task StartupAndDueChecks_DownloadOnlyWhenDue(CatalogAutoUpdateMode mode)
    {
        using var fixture = new Fixture(mode);
        await fixture.Service.InitializeAsync();
        Assert.Equal(mode == CatalogAutoUpdateMode.Disabled ? 0 : 2, fixture.Http.Calls);
        var checkedAt = fixture.Service.Metadata.LastCheckedUtc;
        fixture.Clock.Now += TimeSpan.FromHours(1);
        await fixture.Service.TryCheckAndUpdateAsync();
        Assert.Equal(checkedAt, fixture.Service.Metadata.LastCheckedUtc);
        Assert.Equal(mode == CatalogAutoUpdateMode.Disabled ? 0 : 2, fixture.Http.Calls);
        fixture.Clock.Now += TimeSpan.FromDays(mode == CatalogAutoUpdateMode.Monthly ? 30 : 1);
        await fixture.Service.TryCheckAndUpdateAsync();
        Assert.Equal(mode == CatalogAutoUpdateMode.Disabled ? 0 : 4, fixture.Http.Calls);
        await fixture.Service.TryCheckAndUpdateAsync(forceDownload: true);
        Assert.Equal(mode == CatalogAutoUpdateMode.Disabled ? 2 : 6, fixture.Http.Calls);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FailedOrCancelledDownload_PreservesCatalogAndReschedules(bool cancel)
    {
        using var fixture = new Fixture(CatalogAutoUpdateMode.Daily);
        await fixture.Service.InitializeAsync();
        var original = await File.ReadAllTextAsync(fixture.Store.CatalogFilePath);
        fixture.Clock.Now += TimeSpan.FromDays(1);
        using var cts = new CancellationTokenSource();
        fixture.Http.Failure = cancel
            ? () => { cts.Cancel(); throw new OperationCanceledException(cts.Token); }
            : () => throw new HttpRequestException("Offline access denied");
        await fixture.Service.TryCheckAndUpdateAsync(cancellationToken: cts.Token);
        Assert.Equal(original, await File.ReadAllTextAsync(fixture.Store.CatalogFilePath));
        Assert.True(fixture.Resolver.IsLoaded);
        Assert.False(fixture.Service.IsUpdating);
        Assert.Equal(TimeSpan.FromHours(24), fixture.Clock.LastDue);
        fixture.Http.Failure = null;
        fixture.Clock.Now += TimeSpan.FromDays(1);
        var completion = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        fixture.Service.CatalogUpdated += (_, args) => { if (args.IsSuccess) completion.TrySetResult(); };
        fixture.Clock.Fire();
        await completion.Task.WaitAsync(TimeSpan.FromSeconds(3));
        Assert.Null(fixture.Service.Metadata.LastError);
        Assert.Equal(6, fixture.Http.Calls);
    }

    private sealed class Fixture : IDisposable
    {
        private readonly string _path = Path.Combine(Path.GetTempPath(), "OnlyExo365.Tests", Guid.NewGuid().ToString("N"));
        public FakeClock Clock { get; } = new();
        public FakeHttp Http { get; } = new();
        public PresentationSkuNameResolver Resolver { get; } = new();
        public LicenseCatalogFileStore Store { get; }
        public LicenseCatalogUpdateService Service { get; }
        private readonly HttpClient _client;
        private readonly LicenseCatalogDownloader _downloader;
        public Fixture(CatalogAutoUpdateMode mode)
        {
            var config = new LicenseCatalogConfiguration { LocalCachePath = _path, AutoUpdateMode = mode };
            Store = new(config);
            _client = new HttpClient(Http);
            _downloader = new LicenseCatalogDownloader(config, _client);
            Service = new(config, Store, _downloader, Resolver, Clock);
        }
        public void Dispose() { Service.Dispose(); _downloader.Dispose(); _client.Dispose(); if (Directory.Exists(_path)) Directory.Delete(_path, true); }
    }

    private sealed class FakeHttp : HttpMessageHandler
    {
        public int Calls { get; private set; }
        public Action? Failure { get; set; }
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            Calls++;
            if (request.RequestUri!.Host != "download.microsoft.com")
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent("https://download.microsoft.com/download/offline/licensing.csv") });
            Failure?.Invoke();
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(
                "Product_Display_Name,String_Id,GUID,Service_Plan_Name,Service_Plan_Id,Service_Plans_Included_Friendly_Names\nProduct,SKU_TEST,00000000-0000-0000-0000-000000000001,PLAN,plan-id,Plan") });
        }
    }

    private sealed class FakeClock : TimeProvider
    {
        public DateTimeOffset Now { get; set; } = new(2030, 1, 1, 0, 0, 0, TimeSpan.Zero);
        public TimeSpan LastDue { get; private set; }
        private TimerCallback? _callback;
        private object? _state;
        public override DateTimeOffset GetUtcNow() => Now;
        public override ITimer CreateTimer(TimerCallback callback, object? state, TimeSpan dueTime, TimeSpan period)
        { _callback = callback; _state = state; LastDue = dueTime; return new FakeTimer(); }
        public void Fire() => _callback!(_state);
        private sealed class FakeTimer : ITimer
        {
            public bool Change(TimeSpan dueTime, TimeSpan period) => true;
            public void Dispose() { }
            public ValueTask DisposeAsync() => ValueTask.CompletedTask;
        }
    }
}
