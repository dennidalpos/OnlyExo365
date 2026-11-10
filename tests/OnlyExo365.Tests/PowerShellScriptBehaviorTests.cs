using System.Diagnostics;

namespace OnlyExo365.Tests;

public sealed class PowerShellScriptBehaviorTests
{
    [Theory]
    [InlineData("clean-preservation")]
    [InlineData("clean-locked-file")]
    [InlineData("clean-unsafe-target")]
    [InlineData("clean-reparse-point")]
    [InlineData("clean-tracked-output")]
    [InlineData("clean-native-failure")]
    [InlineData("backup-roundtrip")]
    [InlineData("backup-partial-failure")]
    [InlineData("backup-failed-recovery")]
    [InlineData("backup-missing-entry")]
    [InlineData("owned-process-cleanup")]
    [InlineData("tenant-policy")]
    [InlineData("tenant-minimal-config")]
    [InlineData("tenant-compliance-auth")]
    [InlineData("tenant-failure-report")]
    [InlineData("tenant-app-only-coverage")]
    [InlineData("signing-owned-certificate")]
    [InlineData("installed-payload-signature")]
    [InlineData("catalog-single-sku")]
    [InlineData("catalog-invalid-source")]
    [InlineData("reproducibility-evidence")]
    public async Task OfflineScriptContracts(string scenario)
    {
        var root = TestPathHelper.GetRepositoryPath(".");
        var scratch = Path.Combine(root, "artifacts", "powershell-regressions", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(scratch);
        try
        {
            var start = new ProcessStartInfo("pwsh")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                WorkingDirectory = root
            };
            foreach (var arg in new[] { "-NoLogo", "-NoProfile", "-File", Path.Combine(root, "tests", "powershell", "script-regressions.ps1"), "-RepositoryRoot", root, "-ScratchRoot", scratch, "-Scenario", scenario })
                start.ArgumentList.Add(arg);
            using var process = Process.Start(start)!;
            var output = process.StandardOutput.ReadToEndAsync();
            var error = process.StandardError.ReadToEndAsync();
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(90));
            try
            {
                await process.WaitForExitAsync(timeout.Token);
            }
            catch (OperationCanceledException)
            {
                process.Kill(entireProcessTree: true);
                throw;
            }
            var stdout = await output;
            Assert.True(process.ExitCode == 0, $"{scenario}: {stdout}\n{await error}");
            Assert.Contains($"PASS: {scenario}", stdout, StringComparison.Ordinal);
        }
        finally
        {
            var options = new EnumerationOptions { RecurseSubdirectories = true, AttributesToSkip = FileAttributes.ReparsePoint };
            foreach (var file in Directory.EnumerateFiles(scratch, "*", options))
                File.SetAttributes(file, File.GetAttributes(file) & ~FileAttributes.ReadOnly);
            Directory.Delete(scratch, recursive: true);
        }
    }
}
