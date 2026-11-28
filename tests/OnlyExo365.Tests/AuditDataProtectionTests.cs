namespace OnlyExo365.Tests;

public sealed class AuditDataProtectionTests
{
    [Fact]
    public void Gate_HasNoPersonalDataCleanup()
    {
        var script = File.ReadAllText(TestPathHelper.GetRepositoryPath("scripts", "gate.ps1"));
        Assert.DoesNotContain("Remove-OnlyExo365LocalAppData", script, StringComparison.Ordinal);
        Assert.DoesNotContain("KeepLocalAppData", script, StringComparison.Ordinal);
        Assert.DoesNotContain("CleanPerMachineAppSettings", script, StringComparison.Ordinal);
    }

    [Fact]
    public void Uninstaller_PreservesExportsAndUntrackedInstallationFiles()
    {
        var script = File.ReadAllText(TestPathHelper.GetRepositoryPath("installer", "OnlyExo365.iss"));
        Assert.DoesNotContain("Type: filesandordirs; Name: \"{app}\"", script, StringComparison.Ordinal);
        Assert.DoesNotContain("Type: filesandordirs; Name: \"{localappdata}\\{#ProductName}\\exports\"", script, StringComparison.Ordinal);
    }
}
