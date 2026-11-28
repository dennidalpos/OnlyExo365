using System.Management.Automation.Language;
using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using OnlyExo365.Contracts;
using OnlyExo365.Shell.Services;
using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Tests;

public sealed class AuditRemediationRegressionTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CancelledAuxiliaryProcess_IsTerminated(bool bootstrap)
    {
        var start = new ProcessStartInfo("pwsh") { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true };
        foreach (var argument in new[] { "-NoProfile", "-Command", "Start-Sleep -Seconds 60" }) start.ArgumentList.Add(argument);
        using var process = Process.Start(start)!;
        try
        {
            using var cancellation = new CancellationTokenSource(TimeSpan.FromMilliseconds(150));
            var waiting = bootstrap
                ? InteractiveExchangeBootstrapService.WaitForBootstrapExitAsync(process, cancellation.Token)
                : OnlyExo365.Worker.Program.ReadExecutionPolicyAsync(process, cancellation.Token);
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => waiting.WaitAsync(TimeSpan.FromSeconds(5)));
            Assert.True(process.HasExited);
        }
        finally { if (!process.HasExited) process.Kill(entireProcessTree: true); }
    }

    [Fact]
    public async Task ModuleFallbackDownload_CancelsBlockedHttpWithoutTouchingUserModules()
    {
        using var fixture = new LocalPowerShellFixture();
        var root = Path.Combine(Path.GetTempPath(), "OnlyExo365.Tests", Guid.NewGuid().ToString("N"));
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(8));
        using var cancellation = new CancellationTokenSource();
        try
        {
            var script = ExoSupportCommands.BuildModuleInstallScript(PowerShellModuleBootstrapPolicy.Resolve("ExchangeOnlineManagement")!);
            var ast = Parser.ParseInput(script, out _, out var errors);
            Assert.Empty(errors);
            await fixture.DefineCommandsAsync(string.Join("\n", ast.FindAll(node => node is FunctionDefinitionAst, false).Select(node => node.Extent.Text)));
            await fixture.DefineCommandsAsync($"function Get-CurrentUserModulePath {{ '{root.Replace("'", "''")}' }}");
            var port = ((IPEndPoint)listener.LocalEndpoint).Port;
            var operation = fixture.Engine.ExecuteAsync($"param($DownloadCancellationToken) Install-ModuleFromGalleryPackage -Name Offline.Module -RequiredVersion 1.0.0 -RepositorySourceLocation 'http://127.0.0.1:{port}'",
                new() { ["DownloadCancellationToken"] = cancellation.Token }, cancellationToken: cancellation.Token);
            using var peer = await listener.AcceptTcpClientAsync(timeout.Token);
            cancellation.Cancel();
            var result = await operation.WaitAsync(TimeSpan.FromSeconds(3));
            Assert.True(result.WasCancelled);
            Assert.False(Directory.Exists(root));
        }
        finally
        {
            cancellation.Cancel();
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task CancelledInitialization_DoesNotOpenRunspace()
    {
        using var engine = new PowerShellEngine();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => engine.InitializeAsync(new CancellationToken(true)));
        Assert.False(engine.IsInitialized);
        Assert.Null(engine.Runspace);
    }

    [Fact]
    public async Task LifecycleInvocation_StopsWhenCancelled()
    {
        using var fixture = new LocalPowerShellFixture();
        using var ps = System.Management.Automation.PowerShell.Create();
        ps.Runspace = fixture.Engine.Runspace;
        ps.AddScript("Start-Sleep -Seconds 60");
        using var cancellation = new CancellationTokenSource(TimeSpan.FromMilliseconds(150));
        var invocation = RunspaceLifecycleManager.InvokeAsync(ps, cancellation.Token);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => invocation.WaitAsync(TimeSpan.FromSeconds(10)));
    }

    [Fact]
    public async Task Recovery_WaitsForExecutionLockAndHonorsCancellation()
    {
        using var fixture = new LocalPowerShellFixture();
        var original = fixture.Engine.Runspace!;
        original.Close();
        await fixture.Engine.ExecutionLock.WaitAsync();
        try
        {
            using var cancellation = new CancellationTokenSource();
            var pending = fixture.Engine.ExecuteAsync("'must not execute'", cancellationToken: cancellation.Token);
            Assert.Same(original, fixture.Engine.Runspace);
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => pending);
            Assert.Same(original, fixture.Engine.Runspace);
        }
        finally { fixture.Engine.ExecutionLock.Release(); }
    }

    [Fact]
    public async Task RunspaceRecovery_ClearsAllConnectionState()
    {
        using var fixture = new LocalPowerShellFixture();
        fixture.Engine.Connected = true;
        fixture.Engine.GraphConnected = true;
        fixture.Engine.ComplianceConnected = true;
        fixture.Engine.ConnectedGraphScopesValue = ["LicenseAssignment.ReadWrite.All"];
        fixture.Engine.Runspace!.Close();
        var result = await fixture.Engine.ExecuteAsync("'recovered'");
        Assert.True(result.Success, result.ErrorMessage);
        Assert.False(fixture.Engine.IsConnected);
        Assert.False(fixture.Engine.IsGraphConnected);
        Assert.False(fixture.Engine.IsComplianceConnected);
        Assert.Empty(fixture.Engine.ConnectedGraphScopes);
    }

    [Theory]
    [InlineData("unique", true)]
    [InlineData("missing", false)]
    [InlineData("ambiguous", false)]
    [InlineData("expired", false)]
    [InlineData("future", false)]
    [InlineData("no-key", false)]
    [InlineData("different-subject", false)]
    [InlineData("renewed", true)]
    [InlineData("thumbprint", true)]
    public async Task ComplianceCertificate_BothEntryPointsSelectOnlyOneUsableCertificate(string scenario, bool succeeds)
    {
        using var fixture = new LocalPowerShellFixture();
        await fixture.DefineCommandsAsync("""
            $global:certificateCalls = 0
            function Get-ChildItem {
                [CmdletBinding()] param($LiteralPath)
                if ($LiteralPath -ne 'Cert:\CurrentUser\My') { throw 'Unexpected store.' }
                if ($global:scenario -eq 'thumbprint') { throw 'Explicit thumbprint must not enumerate certificates.' }
                $now = [DateTime]::Now
                $valid = [pscustomobject]@{ Subject="CN=O'Brien"; Thumbprint='UNIQUE'; HasPrivateKey=$true; NotBefore=$now.AddDays(-1); NotAfter=$now.AddDays(1) }
                switch ($global:scenario) {
                    'unique' { $valid }
                    'ambiguous' { $valid; $valid }
                    'expired' { $valid.NotAfter=$now.AddDays(-1); $valid }
                    'future' { $valid.NotBefore=$now.AddDays(1); $valid }
                    'no-key' { $valid.HasPrivateKey=$false; $valid }
                    'different-subject' { $valid.Subject="CN=O'Brien other"; $valid }
                    'renewed' { $valid; [pscustomobject]@{ Subject="CN=O'Brien"; Thumbprint='OLD'; HasPrivateKey=$true; NotBefore=$now.AddDays(-2); NotAfter=$now.AddDays(-1) } }
                }
            }
            function Connect-IPPSSession {
                [CmdletBinding()] param($CertificateThumbprint,$AppId,$Organization,$ConnectionUri,$AzureADAuthorizationEndpointUri,[bool]$ShowBanner)
                $global:certificateCalls++
                $CertificateThumbprint
            }
            """);
        await LoadTenantFunctionsAsync(fixture);
        await fixture.DefineCommandsAsync($"$global:scenario = '{scenario}'");
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.AppCertificate,
            ApplicationId = "offline-app", ExchangeOrganization = "example.onmicrosoft.com",
            CertificateSubjectName = "CN=O'Brien", CertificateThumbprint = scenario == "thumbprint" ? "EXPLICIT" : null
        };
        var worker = await fixture.Engine.ExecuteAsync(ComplianceCommandBuilder.BuildConnectComplianceCommand(configuration));
        Assert.Equal(succeeds, worker.Success);
        var tenant = await fixture.Engine.ExecuteAsync($$"""
            $config = @{ ExchangeEnvironmentName='O365Default'; ApplicationId='offline-app'; ExchangeOrganization='example.onmicrosoft.com';
                CertificateSubjectName="CN=O'Brien"; CertificateThumbprint='{{configuration.CertificateThumbprint}}'; DelegatedOrganization='' }
            $inv = Get-ComplianceConnectCommand $config 'AppCertificate'; $a = $inv.Arguments; & $inv.CommandName @a
            """);
        Assert.Equal(succeeds, tenant.Success);
        if (succeeds)
        {
            var expected = scenario == "thumbprint" ? "EXPLICIT" : "UNIQUE";
            Assert.Equal(expected, Assert.Single(worker.Output).BaseObject.ToString());
            Assert.Equal(expected, Assert.Single(tenant.Output).BaseObject.ToString());
        }
        var calls = await fixture.Engine.ExecuteAsync("$global:certificateCalls");
        Assert.Equal(succeeds ? 2 : 0, Assert.Single(calls.Output).BaseObject);
    }

    [Theory]
    [InlineData("v1.0.0", true)]
    [InlineData("release/1.0", true)]
    [InlineData("v1'quoted", true)]
    [InlineData("v1$(throw)", true)]
    [InlineData("v1;throw", true)]
    [InlineData("", false)]
    [InlineData("-option", false)]
    [InlineData("v1\nvalue=bad", false)]
    [InlineData("v1\rvalue=bad", false)]
    [InlineData("v1'; throw 'injected", false)]
    [InlineData("../bad", false)]
    public async Task ReleaseTag_ValidatesGitSyntaxWithoutEvaluatingInput(string tag, bool valid)
    {
        using var fixture = new LocalPowerShellFixture();
        await fixture.DefineCommandsAsync(File.ReadAllText(TestPathHelper.GetRepositoryPath("build/helpers/release-tag.ps1")));
        var result = await fixture.Engine.ExecuteAsync("param($Tag) Resolve-ReleaseTag -EventName workflow_dispatch -RequestedTag $Tag",
            new() { ["Tag"] = tag });
        Assert.Equal(valid, result.Success);
        if (valid) Assert.Equal(tag, Assert.Single(result.Output).BaseObject.ToString());
    }

    [Fact]
    public async Task ReleaseEvent_UsesPublishedTagAndWorkflowKeepsValuesOutOfScriptSource()
    {
        using var fixture = new LocalPowerShellFixture();
        await fixture.DefineCommandsAsync(File.ReadAllText(TestPathHelper.GetRepositoryPath("build/helpers/release-tag.ps1")));
        var result = await fixture.Engine.ExecuteAsync("Resolve-ReleaseTag -EventName release -PublishedTag v1.0 -RequestedTag ignored");
        Assert.Equal("v1.0", Assert.Single(result.Output).BaseObject.ToString());
        var workflow = File.ReadAllText(TestPathHelper.GetRepositoryPath(".github/workflows/release-signed.yml"));
        foreach (var line in workflow.Split('\n').Where(line => line.Contains("${{", StringComparison.Ordinal)))
            Assert.DoesNotContain("\"${{", line, StringComparison.Ordinal);
        Assert.Contains("-ReleaseTag \"$env:RELEASE_TAG\"", workflow, StringComparison.Ordinal);
        Assert.Contains("gh release upload \"$env:RELEASE_TAG\"", workflow, StringComparison.Ordinal);
    }

    private static async Task LoadTenantFunctionsAsync(LocalPowerShellFixture fixture)
    {
        var ast = Parser.ParseFile(TestPathHelper.GetRepositoryPath("build/run-tenant-validation.ps1"), out _, out var errors);
        Assert.Empty(errors);
        var helperPath = TestPathHelper.GetRepositoryPath("build/helpers/cloud-routing.ps1");
        var helper = File.ReadAllText(helperPath).Replace("$PSScriptRoot", "'" + Path.GetDirectoryName(helperPath)!.Replace("'", "''") + "'");
        await fixture.DefineCommandsAsync(helper + "\n" + string.Join("\n", ast.FindAll(node => node is FunctionDefinitionAst, false).Select(node => node.Extent.Text)));
    }
}
