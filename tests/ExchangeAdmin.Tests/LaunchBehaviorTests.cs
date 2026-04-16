using System.Xml.Linq;

namespace ExchangeAdmin.Tests;

public sealed class LaunchBehaviorTests
{
    [Fact]
    public void PresentationProject_UsesWindowsSubsystem()
    {
        var projectPath = TestPathHelper.GetRepositoryPath(
            "src",
            "ExchangeAdmin.Presentation",
            "ExchangeAdmin.Presentation.csproj");

        var project = XDocument.Load(projectPath);

        Assert.Equal(
            "WinExe",
            project.Root?
                .Elements("PropertyGroup")
                .Elements("OutputType")
                .Single()
                .Value);
    }

    [Fact]
    public void WorkerSupervisor_StartsWorkerWithoutVisibleConsoleWindow()
    {
        var supervisorPath = TestPathHelper.GetRepositoryPath(
            "src",
            "ExchangeAdmin.Infrastructure",
            "Ipc",
            "WorkerSupervisor.cs");

        var source = File.ReadAllText(supervisorPath);

        Assert.Contains("UseShellExecute = false", source, StringComparison.Ordinal);
        Assert.Contains("CreateNoWindow = true", source, StringComparison.Ordinal);
        Assert.Contains("WindowStyle = ProcessWindowStyle.Hidden", source, StringComparison.Ordinal);
        Assert.DoesNotContain("WindowStyle = ProcessWindowStyle.Normal", source, StringComparison.Ordinal);
    }
}
