using System.Text.Json;

namespace OnlyExo365.Tests;

public sealed class BuildScriptSecurityTests
{
    [Fact]
    public void ToolManifest_PinsGitleaksWindowsX64Checksum()
    {
        var manifestPath = GetRepositoryFilePath("build", "tool-manifest.json");

        using var document = JsonDocument.Parse(File.ReadAllText(manifestPath));
        var package = document.RootElement
            .GetProperty("gitleaks")
            .GetProperty("8.30.0")
            .GetProperty("windows_x64");

        Assert.Equal("gitleaks_8.30.0_windows_x64.zip", package.GetProperty("archiveName").GetString());
        Assert.Equal("54fe94f644b832dd08e8c3a5915efb3bfa862386d59fb27ca0792cb687a83573", package.GetProperty("sha256").GetString());
    }

    [Fact]
    public void SigningScripts_UseCanonicalTimestampAuthority_AndExeOnlyTargets()
    {
        var helperPath = GetRepositoryFilePath("build", "signing-helpers.ps1");
        var helperScript = File.ReadAllText(helperPath);

        Assert.Contains("TimestampUrl must use HTTP or HTTPS", helperScript, StringComparison.Ordinal);
        Assert.Contains("Get-ChildItem -Path $entry.Path -Recurse -File -Include *.exe", helperScript, StringComparison.Ordinal);
        Assert.DoesNotContain("*.msi", helperScript, StringComparison.Ordinal);

        foreach (var relativePath in new[]
                 {
                     new[] { "build", "sign-artifacts.ps1" },
                     new[] { "build", "validate-artifact-signing.ps1" }
                 })
        {
            var scriptPath = GetRepositoryFilePath(relativePath[0], relativePath[1]);
            var script = File.ReadAllText(scriptPath);

            Assert.Contains("http://timestamp.digicert.com", script, StringComparison.Ordinal);
            Assert.Contains("signing-helpers.ps1", script, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void ValidateArtifactSigningScript_CleansDisposableCertificatesBySubjectBeforeAndAfterExecution()
    {
        var scriptPath = GetRepositoryFilePath("build", "validate-artifact-signing.ps1");
        var script = File.ReadAllText(scriptPath);

        Assert.Contains("function Remove-CertificatesBySubject", script, StringComparison.Ordinal);
        Assert.DoesNotContain("function Add-CertificateToStoreSilently", script, StringComparison.Ordinal);
        Assert.DoesNotContain("CurrentUser\\Root", script, StringComparison.Ordinal);
        Assert.DoesNotContain("CurrentUser\\TrustedPublisher", script, StringComparison.Ordinal);
        Assert.Contains("-AllowUntrustedSigner", script, StringComparison.Ordinal);
        Assert.Contains("Remove-CertificatesBySubject -StorePath \"Cert:\\CurrentUser\\My\" -Subject $SignerSubject -DeleteKey", script, StringComparison.Ordinal);
    }

    [Fact]
    public void SecretScanScript_UsesPinnedManifestAndHashValidationForDownloads()
    {
        var scriptPath = GetRepositoryFilePath("build", "run-secret-scan.ps1");
        var script = File.ReadAllText(scriptPath);

        Assert.Contains("scripts\\internal\\common.ps1", script, StringComparison.Ordinal);
        Assert.Contains("tool-manifest.json", script, StringComparison.Ordinal);
        Assert.Contains("Get-FileHash -Algorithm SHA256", script, StringComparison.Ordinal);
        Assert.Contains("archiveSha256", script, StringComparison.Ordinal);
        Assert.Contains("Downloaded gitleaks archive hash mismatch", script, StringComparison.Ordinal);
        Assert.DoesNotContain("function Write-Step", script, StringComparison.Ordinal);
    }

    [Fact]
    public void NuGetVulnerabilityScanScript_RestoresSolutionAndScansProjectsIndividually()
    {
        var scriptPath = GetRepositoryFilePath("build", "assert-no-vulnerable-packages.ps1");
        var script = File.ReadAllText(scriptPath);

        Assert.Contains("scripts\\internal\\common.ps1", script, StringComparison.Ordinal);
        Assert.Contains("& dotnet sln $SolutionPath list", script, StringComparison.Ordinal);
        Assert.Contains("Invoke-SolutionRestoreForPackageInspection", script, StringComparison.Ordinal);
        Assert.Contains("foreach ($projectPath in $projectPaths)", script, StringComparison.Ordinal);
        Assert.Contains("\"No known NuGet vulnerabilities reported across $($projectPaths.Count) project(s).\"", script, StringComparison.Ordinal);
        Assert.DoesNotContain("function Write-Step", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildScript_StopsArtifactProcessesBeforeRetryingLockedArtifactsCleanup()
    {
        var scriptPath = GetRepositoryFilePath("build", "build.ps1");
        var script = File.ReadAllText(scriptPath);

        Assert.Contains("function Get-ArtifactOnlyExo365Processes", script, StringComparison.Ordinal);
        Assert.Contains("function Stop-ArtifactOnlyExo365Processes", script, StringComparison.Ordinal);
        Assert.Contains("Detected running OnlyExo365 process(es) from artifacts", script, StringComparison.Ordinal);
        Assert.Contains("Stop-Process -Id $process.ProcessId -Force -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("Remove-DirectoryRobust -Path $OutputDir -Description \"artifacts output\" -ArtifactRootForProcessStop $OutputDir", script, StringComparison.Ordinal);
        Assert.DoesNotContain("[switch]$Msi", script, StringComparison.Ordinal);
        Assert.DoesNotContain("helpers\\msi-packaging.ps1", script, StringComparison.Ordinal);
    }

    [Fact]
    public void CommonAndDoctorScripts_ResolveInnoSetupCompilerInsteadOfWixOrIexpress()
    {
        var commonPath = GetRepositoryFilePath("scripts", "internal", "common.ps1");
        var doctorPath = GetRepositoryFilePath("scripts", "agents", "doctor.ps1");
        var commonScript = File.ReadAllText(commonPath);
        var doctorScript = File.ReadAllText(doctorPath);

        Assert.Contains("function Get-InnoSetupCandidateDirectories", commonScript, StringComparison.Ordinal);
        Assert.Contains("function Get-InnoSetupCompilerPath", commonScript, StringComparison.Ordinal);
        Assert.Contains("INNOSETUP_BIN", commonScript, StringComparison.Ordinal);
        Assert.DoesNotContain("wix314-binaries", commonScript, StringComparison.Ordinal);
        Assert.DoesNotContain("WiX Toolset", commonScript, StringComparison.Ordinal);

        Assert.Contains("Get-InnoSetupBinPath -RepositoryRoot $repositoryRoot", doctorScript, StringComparison.Ordinal);
        Assert.Contains("Get-InnoSetupCompilerPath -RepositoryRoot $repositoryRoot", doctorScript, StringComparison.Ordinal);
        Assert.DoesNotContain("iexpress.exe", doctorScript, StringComparison.Ordinal);
        Assert.DoesNotContain("WiX", doctorScript, StringComparison.Ordinal);
    }

    [Fact]
    public void InstallInnoSetupScript_VerifiesExistingCompilerAndUsesSupportedPackageManagersOnly()
    {
        var scriptPath = GetRepositoryFilePath("scripts", "Install-InnoSetup.ps1");
        var script = File.ReadAllText(scriptPath);

        Assert.Contains("internal/common.ps1", script, StringComparison.Ordinal);
        Assert.Contains("function Test-InnoSetupAvailable", script, StringComparison.Ordinal);
        Assert.Contains("Get-InnoSetupCompilerPath -RepositoryRoot $RepositoryRoot", script, StringComparison.Ordinal);
        Assert.Contains("[switch]$Install", script, StringComparison.Ordinal);
        Assert.Contains("[ValidateSet(\"Auto\", \"Winget\", \"Chocolatey\")]", script, StringComparison.Ordinal);
        Assert.Contains("--id\", \"JRSoftware.InnoSetup\"", script, StringComparison.Ordinal);
        Assert.Contains("\"innosetup\"", script, StringComparison.Ordinal);
        Assert.Contains("--version=6.7.1", script, StringComparison.Ordinal);
        Assert.DoesNotContain("WiX", script, StringComparison.Ordinal);
        Assert.DoesNotContain("msiexec", script, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PackAndPublishScripts_AreExeOnly()
    {
        var packPath = GetRepositoryFilePath("scripts", "pack.ps1");
        var publishPath = GetRepositoryFilePath("scripts", "agents", "publish.ps1");
        var packScript = File.ReadAllText(packPath);
        var publishScript = File.ReadAllText(publishPath);

        Assert.Contains("-PublishPath", packScript, StringComparison.Ordinal);
        Assert.Contains("OnlyExo365.Setup.exe", packScript, StringComparison.Ordinal);
        Assert.DoesNotContain("OnlyExo365.msi", packScript, StringComparison.Ordinal);

        Assert.Contains("[string]$SetupExePath = \"artifacts/packages/OnlyExo365.Setup.exe\"", publishScript, StringComparison.Ordinal);
        Assert.DoesNotContain("[string]$MsiPath", publishScript, StringComparison.Ordinal);
        Assert.DoesNotContain("MSI package not found", publishScript, StringComparison.Ordinal);
    }

    [Fact]
    public void LegacyCreateMsiEntryPoint_HasBeenRemovedFromRepository()
    {
        var legacyPath = GetRepositoryFilePath("build", "create-msi.ps1");
        var wixAuthoringPath = GetRepositoryFilePath("installer", "OnlyExo365.wxs");

        Assert.False(File.Exists(legacyPath));
        Assert.False(File.Exists(wixAuthoringPath));
    }

    [Fact]
    public void CleanScript_RecursivelyRemovesRepositoryBinObjDirectoriesOutsideArtifacts()
    {
        var scriptPath = GetRepositoryFilePath("build", "clean.ps1");
        var script = File.ReadAllText(scriptPath);

        Assert.Contains("Get-ChildItem -Path $SolutionDir -Directory -Recurse -Force", script, StringComparison.Ordinal);
        Assert.Contains("$_.Name -in @('bin', 'obj')", script, StringComparison.Ordinal);
        Assert.Contains("$_.FullName -notlike \"$ArtifactsDir\\*\"", script, StringComparison.Ordinal);
        Assert.Contains("$_.FullName -notlike \"$SolutionDir\\.git\\*\"", script, StringComparison.Ordinal);
        Assert.Contains("$measurement.PSObject.Properties['Sum']", script, StringComparison.Ordinal);
        Assert.Contains("Write-Success \"Cleaned $cleanedBinObjDirectories bin/obj directorie(s)\"", script, StringComparison.Ordinal);
    }

    [Fact]
    public void GateScript_CleansLocalAppStateAndRunsCanonicalArtifactBuild()
    {
        var scriptPath = GetRepositoryFilePath("scripts", "gate.ps1");
        var script = File.ReadAllText(scriptPath);

        Assert.Contains("scripts\\clean.ps1", script, StringComparison.Ordinal);
        Assert.Contains("SpecialFolder]::LocalApplicationData", script, StringComparison.Ordinal);
        Assert.Contains("SpecialFolder]::ApplicationData", script, StringComparison.Ordinal);
        Assert.Contains("SpecialFolder]::CommonApplicationData", script, StringComparison.Ordinal);
        Assert.Contains("[switch]$CleanPerMachineAppSettings", script, StringComparison.Ordinal);
        Assert.Contains("scripts\\Install-InnoSetup.ps1", script, StringComparison.Ordinal);
        Assert.Contains("scripts\\agents\\doctor.ps1", script, StringComparison.Ordinal);
        Assert.Contains("scripts\\agents\\compile.ps1", script, StringComparison.Ordinal);
        Assert.Contains("scripts\\agents\\test.ps1", script, StringComparison.Ordinal);
        Assert.Contains("build\\assert-no-vulnerable-packages.ps1", script, StringComparison.Ordinal);
        Assert.Contains("build\\run-secret-scan.ps1", script, StringComparison.Ordinal);
        Assert.Contains("scripts\\pack.ps1", script, StringComparison.Ordinal);
        Assert.Contains("OnlyExo365.Setup.exe", script, StringComparison.Ordinal);
        Assert.DoesNotContain("OnlyExo365.msi", script, StringComparison.Ordinal);
    }

    [Fact]
    public void InnoInstallerAuthoring_KeepsPerMachineShortcutsAndUninstallCleanup()
    {
        var scriptPath = GetRepositoryFilePath("installer", "OnlyExo365.iss");
        var script = File.ReadAllText(scriptPath);

        Assert.Contains("AppId={{B9E5A61C-8D6A-4B10-8F50-2BB72D7497F3}", script, StringComparison.Ordinal);
        Assert.Contains("DefaultDirName={autopf}\\{#ProductName}", script, StringComparison.Ordinal);
        Assert.Contains("PrivilegesRequired=admin", script, StringComparison.Ordinal);
        Assert.Contains("ArchitecturesInstallIn64BitMode=x64compatible", script, StringComparison.Ordinal);
        Assert.Contains("ArchitecturesAllowed=x64compatible", script, StringComparison.Ordinal);
        Assert.Contains("LicenseFile={#LicensePath}", script, StringComparison.Ordinal);
        Assert.Contains("Source: \"{#PublishDirX64}\\*\"", script, StringComparison.Ordinal);
        Assert.DoesNotContain("PublishDirX86", script, StringComparison.Ordinal);
        Assert.Contains("Check: Is64BitInstallMode", script, StringComparison.Ordinal);
        Assert.DoesNotContain("Check: not Is64BitInstallMode", script, StringComparison.Ordinal);
        Assert.Contains("Name: \"{commonprograms}\\{#ProductName}\"", script, StringComparison.Ordinal);
        Assert.Contains("Name: \"{commondesktop}\\{#ProductName}\"", script, StringComparison.Ordinal);
        Assert.Contains("ValueName: \"InstallLocation\"", script, StringComparison.Ordinal);
        Assert.Contains("ValueName: \"LogDirectory\"", script, StringComparison.Ordinal);
        Assert.Contains("ValueName: \"SecretDirectory\"", script, StringComparison.Ordinal);
        Assert.Contains("ValueName: \"ExportDirectory\"", script, StringComparison.Ordinal);
        Assert.DoesNotContain("[UninstallRun]", script, StringComparison.Ordinal);
        Assert.DoesNotContain("OnlyExo365.ServiceCleanup.ps1", script, StringComparison.Ordinal);
        Assert.Contains("[UninstallDelete]", script, StringComparison.Ordinal);
    }

    [Fact]
    public void SmokeTestScripts_BackUpAndRestorePreExistingLocalDataBeforeInstallerValidation()
    {
        var scriptPath = GetRepositoryFilePath("build", "run-smoke-tests.ps1");
        var helperPath = GetRepositoryFilePath("build", "helpers", "run-smoke-tests.helpers.ps1");
        var script = File.ReadAllText(scriptPath);
        var helper = File.ReadAllText(helperPath);
        var combined = string.Concat(script, Environment.NewLine, helper);

        Assert.Contains("[string]$SetupExePath = \"artifacts/packages/OnlyExo365.Setup.exe\"", script, StringComparison.Ordinal);
        Assert.Contains("function Get-UninstallInvocation", helper, StringComparison.Ordinal);
        Assert.Contains("function Resolve-SmokePublishPath", helper, StringComparison.Ordinal);
        Assert.Contains("Setup EXE installation failed", script, StringComparison.Ordinal);
        Assert.Contains("Setup EXE reinstall failed", script, StringComparison.Ordinal);
        Assert.Contains("Setup reinstall preserved a single uninstall registry entry.", script, StringComparison.Ordinal);
        Assert.Contains("Setup uninstall removed the registered installation and residual files.", script, StringComparison.Ordinal);
        Assert.Contains("Moving it to temporary backup before installer uninstall validation", combined, StringComparison.Ordinal);
        Assert.Contains("Restored pre-existing OnlyExo365 local data after installer validation.", combined, StringComparison.Ordinal);
        Assert.Contains("Recovered pre-existing OnlyExo365 local data during final cleanup after an interrupted installer smoke test run.", combined, StringComparison.Ordinal);
        Assert.DoesNotContain("artifacts/packages/OnlyExo365.msi", script, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerShellBootstrapPolicy_PinsApprovedModuleVersions()
    {
        var policyPath = GetRepositoryFilePath("src", "OnlyExo365.Worker", "Data", "PowerShellModuleBootstrapPolicy.json");
        using var document = JsonDocument.Parse(File.ReadAllText(policyPath));

        Assert.Equal("PSGallery", document.RootElement.GetProperty("repositoryName").GetString());
        Assert.Equal("https://www.powershellgallery.com/api/v2", document.RootElement.GetProperty("repositorySourceLocation").GetString());

        var modules = document.RootElement.GetProperty("modules").EnumerateArray().ToArray();
        Assert.Contains(modules, module =>
            module.GetProperty("moduleName").GetString() == "ExchangeOnlineManagement" &&
            module.GetProperty("requiredVersion").GetString() == "3.9.2");
        Assert.Contains(modules, module =>
            module.GetProperty("moduleName").GetString() == "Microsoft.Graph.Authentication" &&
            module.GetProperty("requiredVersion").GetString() == "2.35.1");
        Assert.Contains(modules, module =>
            module.GetProperty("requestName").GetString() == "Microsoft.Graph" &&
            module.GetProperty("requiredModules").EnumerateArray().Select(item => item.GetString()).SequenceEqual(new[]
            {
                "Microsoft.Graph.Authentication",
                "Microsoft.Graph.Users",
                "Microsoft.Graph.Users.Actions",
                "Microsoft.Graph.Identity.DirectoryManagement"
            }));
    }

    [Fact]
    public void ExoSupportCommands_InstallModuleScript_UsesAllowlistAndRequiredVersion()
    {
        var scriptPath = GetRepositoryFilePath("src", "OnlyExo365.Worker", "PowerShell", "ExoSupportCommands.cs");
        var script = File.ReadAllText(scriptPath);

        Assert.Contains("PowerShellModuleBootstrapPolicy.Resolve(request.ModuleName)", script, StringComparison.Ordinal);
        Assert.Contains("Version = $requiredVersion", script, StringComparison.Ordinal);
        Assert.Contains("Install-PSResource", script, StringComparison.Ordinal);
        Assert.Contains("Install-ModuleFromGalleryPackage", script, StringComparison.Ordinal);
        Assert.Contains("$moduleNames =", script, StringComparison.Ordinal);
        Assert.Contains("/package/$Name/$RequiredVersion", script, StringComparison.Ordinal);
        Assert.Contains("System.Net.Http.HttpClient", script, StringComparison.Ordinal);
        Assert.Contains("System.IO.Compression.ZipFile", script, StringComparison.Ordinal);
        Assert.Contains("Join-Path $expandedPath \"\"$Name.psd1\"\"", script, StringComparison.Ordinal);
        Assert.Contains("does not contain a module manifest. Top-level entries:", script, StringComparison.Ordinal);
        Assert.DoesNotContain("Install-PackageProvider -Name NuGet", script, StringComparison.Ordinal);
        Assert.DoesNotContain("Invoke-WebRequest -Uri $downloadUri", script, StringComparison.Ordinal);
        Assert.DoesNotContain("Expand-Archive -Path $packagePath", script, StringComparison.Ordinal);
        Assert.DoesNotContain("Install-Module -Name '{safeModuleName}' -Force -AllowClobber -Scope CurrentUser -Confirm:$false -ErrorAction Stop", script, StringComparison.Ordinal);
    }

    [Fact]
    public void PublishReleaseAssetsScript_PreparesZipChecksumsAndGitHubEnvExports()
    {
        var scriptPath = GetRepositoryFilePath("build", "publish-release-assets.ps1");
        var script = File.ReadAllText(scriptPath);

        Assert.Contains("scripts\\internal\\common.ps1", script, StringComparison.Ordinal);
        Assert.Contains("Compress-Archive", script, StringComparison.Ordinal);
        Assert.Contains("Get-FileHash -Algorithm SHA256", script, StringComparison.Ordinal);
        Assert.Contains("RELEASE_ASSET_ZIP=", script, StringComparison.Ordinal);
        Assert.Contains("RELEASE_ASSET_SETUP_EXE=", script, StringComparison.Ordinal);
        Assert.Contains("RELEASE_ASSET_CHECKSUMS=", script, StringComparison.Ordinal);
        Assert.Contains("RELEASE_ASSET_MANIFEST=", script, StringComparison.Ordinal);
        Assert.Contains("OnlyExo365-$sanitizedTag-$resolvedRuntimeIdentifier-publish.zip", script, StringComparison.Ordinal);
        Assert.Contains("OnlyExo365-$sanitizedTag-$resolvedRuntimeIdentifier-setup.exe", script, StringComparison.Ordinal);
        Assert.DoesNotContain("RELEASE_ASSET_MSI", script, StringComparison.Ordinal);
        Assert.DoesNotContain(".msi", script, StringComparison.Ordinal);
        Assert.DoesNotContain("function Write-Step", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildPerimeterScripts_ReuseSharedCommonLoggingHelpers()
    {
        foreach (var relativePath in new[]
                 {
                     new[] { "build", "assert-no-vulnerable-packages.ps1" },
                     new[] { "build", "run-secret-scan.ps1" },
                     new[] { "build", "publish-release-assets.ps1" },
                     new[] { "build", "signing-helpers.ps1" },
                     new[] { "build", "clean.ps1" },
                     new[] { "build", "run-smoke-tests.ps1" },
                     new[] { "build", "run-tenant-validation.ps1" },
                     new[] { "build", "verify-package-reproducibility.ps1" }
                 })
        {
            var scriptPath = GetRepositoryFilePath(relativePath[0], relativePath[1]);
            var script = File.ReadAllText(scriptPath);

            Assert.Contains("scripts\\internal\\common.ps1", script, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void LegacyManualReleaseBaselineEntryPoint_HasBeenRemovedFromRepository()
    {
        var legacyPath = GetRepositoryFilePath("build", "release-baseline.ps1");

        Assert.False(File.Exists(legacyPath));
    }

    [Fact]
    public void AppIconGeneration_UsesSourceSpecAndGeneratedAssetFolder()
    {
        var scriptPath = GetRepositoryFilePath("build", "generate-app-icon.ps1");
        var setupScriptPath = GetRepositoryFilePath("build", "create-setup-exe.ps1");
        var projectPath = GetRepositoryFilePath("src", "OnlyExo365.Shell", "OnlyExo365.Shell.csproj");
        var sourceSpecPath = GetRepositoryFilePath("src", "OnlyExo365.Shell", "Assets", "Source", "app-icon.spec.json");
        var generatedIcoPath = GetRepositoryFilePath("src", "OnlyExo365.Shell", "Assets", "Generated", "AppIcon.ico");
        var generatedPngPath = GetRepositoryFilePath("src", "OnlyExo365.Shell", "Assets", "Generated", "AppIcon.png");

        var script = File.ReadAllText(scriptPath);
        var setupScript = File.ReadAllText(setupScriptPath);
        var project = File.ReadAllText(projectPath);

        Assert.True(File.Exists(sourceSpecPath));
        Assert.True(File.Exists(generatedIcoPath));
        Assert.True(File.Exists(generatedPngPath));
        Assert.Contains("Source/app-icon.spec.json", script, StringComparison.Ordinal);
        Assert.Contains("Assets/Generated/AppIcon.ico", script, StringComparison.Ordinal);
        Assert.Contains("Assets/Generated/AppIcon.png", script, StringComparison.Ordinal);
        Assert.Contains("ConvertFrom-Json", script, StringComparison.Ordinal);
        Assert.Contains("Assets\\Generated\\AppIcon.ico", setupScript, StringComparison.Ordinal);
        Assert.Contains("<ApplicationIcon>Assets\\Generated\\AppIcon.ico</ApplicationIcon>", project, StringComparison.Ordinal);
    }

    [Fact]
    public void FallbackAndManualInstructionText_RemainsEnglishInBaselinePaths()
    {
        foreach (var relativePath in new[]
                 {
                     new[] { "src", "OnlyExo365.Shell", "ViewModels", "MigrationViewModel.BatchCreation.cs" },
                     new[] { "src", "OnlyExo365.Worker", "PowerShell", "ExoSupportCommands.cs" },
                     new[] { "src", "OnlyExo365.Worker", "PowerShell", "PowerShellModuleBootstrapPolicy.cs" }
                 })
        {
            var content = File.ReadAllText(GetRepositoryFilePath(relativePath));

            Assert.DoesNotContain("Installare", content, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Eseguire", content, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("scaricare", content, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("eseguire", content, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("creato", content, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("completato", content, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void CreateSetupExeScript_UsesInnoSetupCompilerAndInstallerAuthoring()
    {
        var scriptPath = GetRepositoryFilePath("build", "create-setup-exe.ps1");
        var script = File.ReadAllText(scriptPath);

        Assert.Contains("[string]$PublishPath", script, StringComparison.Ordinal);
        Assert.Contains("installer\\OnlyExo365.iss", script, StringComparison.Ordinal);
        Assert.Contains("Get-InnoSetupCompilerPath -RepositoryRoot $repositoryRoot", script, StringComparison.Ordinal);
        Assert.Contains("/DPublishDirX64=$resolvedPublishPathX64", script, StringComparison.Ordinal);
        Assert.Contains("/DLicensePath=$resolvedLicensePath", script, StringComparison.Ordinal);
        Assert.DoesNotContain("PublishDirX86", script, StringComparison.Ordinal);
        Assert.DoesNotContain("CleanupScriptPath", script, StringComparison.Ordinal);
        Assert.DoesNotContain("iexpress.exe", script, StringComparison.Ordinal);
        Assert.DoesNotContain("msiexec.exe", script, StringComparison.Ordinal);
        Assert.DoesNotContain("MsiPath", script, StringComparison.Ordinal);
    }

    [Fact]
    public void CanonicalBuildAndPackScripts_UseSwitchFlagsForBooleanParameters()
    {
        var buildScriptPath = GetRepositoryFilePath("scripts", "build.ps1");
        var canonicalBuildScriptPath = GetRepositoryFilePath("build", "build.ps1");
        var compileScriptPath = GetRepositoryFilePath("scripts", "agents", "compile.ps1");
        var packScriptPath = GetRepositoryFilePath("scripts", "pack.ps1");
        var testScriptPath = GetRepositoryFilePath("scripts", "agents", "test.ps1");
        var buildScript = File.ReadAllText(buildScriptPath);
        var canonicalBuildScript = File.ReadAllText(canonicalBuildScriptPath);
        var compileScript = File.ReadAllText(compileScriptPath);
        var packScript = File.ReadAllText(packScriptPath);
        var testScript = File.ReadAllText(testScriptPath);

        Assert.Contains("[switch]$Clean = $true", buildScript, StringComparison.Ordinal);
        Assert.Contains("[switch]$LockedMode = $false", buildScript, StringComparison.Ordinal);
        Assert.Contains("[switch]$NoRestore = $false", buildScript, StringComparison.Ordinal);
        Assert.Contains("[switch]$SelfContained = $false", buildScript, StringComparison.Ordinal);
        Assert.Contains("\"-LockedMode:$([bool]$LockedMode)\"", buildScript, StringComparison.Ordinal);
        Assert.Contains("\"-NoRestore:$([bool]$NoRestore)\"", buildScript, StringComparison.Ordinal);

        Assert.Contains("[switch]$NoRestore = $false", canonicalBuildScript, StringComparison.Ordinal);
        Assert.Contains("Restore skipped by -NoRestore", canonicalBuildScript, StringComparison.Ordinal);
        Assert.Contains("scripts/build.ps1", compileScript, StringComparison.Ordinal);
        Assert.Contains("\"-NoRestore:$true\"", compileScript, StringComparison.Ordinal);
        Assert.DoesNotContain("Invoke-DotNetCommand -Arguments $arguments -ErrorMessage \"dotnet build failed\"", compileScript, StringComparison.Ordinal);
        Assert.Contains("Invoke-RepositoryBootstrap", compileScript, StringComparison.Ordinal);

        Assert.Contains("[switch]$LockedMode = $true", packScript, StringComparison.Ordinal);
        Assert.Contains("[switch]$Clean = $false", packScript, StringComparison.Ordinal);
        Assert.Contains("[switch]$SelfContained = $false", packScript, StringComparison.Ordinal);
        Assert.Contains("\"-Clean:$([bool]$Clean)\"", packScript, StringComparison.Ordinal);

        Assert.Contains("\"--disable-build-servers\"", testScript, StringComparison.Ordinal);
        Assert.Contains("Invoke-RepositoryBootstrap", testScript, StringComparison.Ordinal);
        Assert.Contains("$PrimaryRuntimeIdentifier = @($ResolvedRuntimeIdentifiers)[0]", canonicalBuildScript, StringComparison.Ordinal);
        Assert.DoesNotContain("$ResolvedRuntimeIdentifiers[0]", canonicalBuildScript, StringComparison.Ordinal);
    }

    [Fact]
    public void CanonicalScripts_EnforcePinnedDotNetSdkAndCanonicalArtifactPaths()
    {
        var commonPath = GetRepositoryFilePath("scripts", "internal", "common.ps1");
        var bootstrapPath = GetRepositoryFilePath("scripts", "bootstrap.ps1");
        var doctorPath = GetRepositoryFilePath("scripts", "agents", "doctor.ps1");
        var packPath = GetRepositoryFilePath("scripts", "pack.ps1");
        var publishPath = GetRepositoryFilePath("scripts", "agents", "publish.ps1");
        var smokePath = GetRepositoryFilePath("build", "run-smoke-tests.ps1");

        var commonScript = File.ReadAllText(commonPath);
        var bootstrapScript = File.ReadAllText(bootstrapPath);
        var doctorScript = File.ReadAllText(doctorPath);
        var packScript = File.ReadAllText(packPath);
        var publishScript = File.ReadAllText(publishPath);
        var smokeScript = File.ReadAllText(smokePath);

        Assert.Contains("function Assert-DotNetSdkPinnedVersion", commonScript, StringComparison.Ordinal);
        Assert.Contains("function Get-InnoSetupCompilerPath", commonScript, StringComparison.Ordinal);
        Assert.Contains("function Invoke-RepositoryBootstrap", commonScript, StringComparison.Ordinal);
        Assert.Contains("scripts/bootstrap.ps1", commonScript, StringComparison.Ordinal);
        Assert.Contains("global.json pins SDK version", commonScript, StringComparison.Ordinal);
        Assert.Contains("Assert-DotNetSdkPinnedVersion -RepositoryRoot $repositoryRoot", bootstrapScript, StringComparison.Ordinal);
        Assert.Contains("Assert-DotNetSdkPinnedVersion -RepositoryRoot $repositoryRoot", doctorScript, StringComparison.Ordinal);
        Assert.Contains("$packagesDirectory = Join-Path $repositoryRoot \"artifacts\\\\packages\"", packScript, StringComparison.Ordinal);
        Assert.Contains("OnlyExo365.Setup.exe", packScript, StringComparison.Ordinal);
        Assert.DoesNotContain("OnlyExo365.msi", packScript, StringComparison.Ordinal);
        Assert.Contains("[string]$RuntimeIdentifier = \"win-x64\"", publishScript, StringComparison.Ordinal);
        Assert.Contains("\"-RuntimeIdentifier\", $RuntimeIdentifier", publishScript, StringComparison.Ordinal);
        Assert.Contains("artifacts/packages/OnlyExo365.Setup.exe", publishScript, StringComparison.Ordinal);
        Assert.DoesNotContain("artifacts/packages/OnlyExo365.msi", publishScript, StringComparison.Ordinal);
        Assert.Contains("artifacts/packages/OnlyExo365.Setup.exe", smokeScript, StringComparison.Ordinal);
    }

    [Fact]
    public void ArchitectureConstraintsScript_GuardsRefactoredBootstrapTopology_AndInnoPackaging()
    {
        var scriptPath = GetRepositoryFilePath("build", "assert-architecture-constraints.ps1");
        var script = File.ReadAllText(scriptPath);

        Assert.Contains("AppModuleCatalog.cs", script, StringComparison.Ordinal);
        Assert.Contains("AppModuleFactory.cs", script, StringComparison.Ordinal);
        Assert.Contains("AppShellModuleRegistrar.cs", script, StringComparison.Ordinal);
        Assert.Contains("AppPageLoaderCatalog.cs", script, StringComparison.Ordinal);
        Assert.Contains("build/create-setup-exe.ps1", script, StringComparison.Ordinal);
        Assert.Contains("installer/OnlyExo365.iss", script, StringComparison.Ordinal);
        Assert.Contains("AppCompositionRoot.cs'; MaxLines = 80", script, StringComparison.Ordinal);
        Assert.Contains("AppModuleFactory.cs'; MaxLines = 120", script, StringComparison.Ordinal);
        Assert.Contains("AppShellModuleRegistrar.cs'; MaxLines = 190", script, StringComparison.Ordinal);
        Assert.Contains("AppPageLoaderCatalog.cs'; MaxLines = 80", script, StringComparison.Ordinal);
    }

    [Fact]
    public void ReleaseSignedWorkflow_UploadsDurableExeOnlyAssetsToGitHubRelease()
    {
        var workflowPath = GetRepositoryFilePath(".github", "workflows", "release-signed.yml");
        var workflow = File.ReadAllText(workflowPath);

        Assert.Contains("contents: write", workflow, StringComparison.Ordinal);
        Assert.Contains("uses: ./.github/actions/windows-release-baseline", workflow, StringComparison.Ordinal);
        Assert.Contains("./scripts/agents/publish.ps1", workflow, StringComparison.Ordinal);
        Assert.Contains("gh release upload", workflow, StringComparison.Ordinal);
        Assert.Contains("RELEASE_ASSET_SETUP_EXE", workflow, StringComparison.Ordinal);
        Assert.Contains("release-assets-signed", workflow, StringComparison.Ordinal);
        Assert.DoesNotContain("RELEASE_ASSET_MSI", workflow, StringComparison.Ordinal);
    }

    [Fact]
    public void ReleaseValidationWorkflow_RunsExeOnlyPackagingAndSmokeValidation()
    {
        var workflowPath = GetRepositoryFilePath(".github", "workflows", "release-validation.yml");
        var workflow = File.ReadAllText(workflowPath);

        Assert.Contains("uses: ./.github/actions/windows-release-baseline", workflow, StringComparison.Ordinal);
        Assert.Contains("./build/run-smoke-tests.ps1 -PublishPath artifacts/publish -SetupExePath artifacts/packages/OnlyExo365.Setup.exe -ReportPath artifacts/smoke/smoke-report.json", workflow, StringComparison.Ordinal);
        Assert.DoesNotContain("OnlyExo365.msi", workflow, StringComparison.Ordinal);
    }

    [Fact]
    public void WindowsReleaseBaselineCompositeAction_RunsCanonicalPreSigningValidationAndPackaging()
    {
        var actionPath = GetRepositoryFilePath(".github", "actions", "windows-release-baseline", "action.yml");
        var action = File.ReadAllText(actionPath);

        Assert.Contains("actions/setup-dotnet@v4", action, StringComparison.Ordinal);
        Assert.Contains("choco install innosetup", action, StringComparison.Ordinal);
        Assert.Contains("./scripts/agents/doctor.ps1 -CheckPackaging", action, StringComparison.Ordinal);
        Assert.Contains("./scripts/bootstrap.ps1 -RuntimeIdentifier ${{ inputs.runtime_identifier }}", action, StringComparison.Ordinal);
        Assert.Contains("./scripts/agents/compile.ps1 -Configuration Debug -RuntimeIdentifier ${{ inputs.runtime_identifier }} -NoBootstrap", action, StringComparison.Ordinal);
        Assert.Contains("./build/assert-no-vulnerable-packages.ps1 -SolutionPath OnlyExo365.sln -ReportPath artifacts/security/nuget-vulnerabilities.json", action, StringComparison.Ordinal);
        Assert.Contains("./build/run-secret-scan.ps1 -SourcePath . -ReportPath artifacts/security/gitleaks.sarif", action, StringComparison.Ordinal);
        Assert.Contains("./scripts/agents/test.ps1 -Configuration Debug -ResultsDirectory artifacts/test-results/unit -NoBootstrap -NoBuild -NoRestore", action, StringComparison.Ordinal);
        Assert.Contains("./build/assert-code-coverage.ps1", action, StringComparison.Ordinal);
        Assert.Contains("./build/assert-architecture-constraints.ps1", action, StringComparison.Ordinal);
        Assert.Contains("./scripts/pack.ps1 -Configuration Release -Clean:$false -LockedMode -RuntimeIdentifier ${{ inputs.runtime_identifier }}", action, StringComparison.Ordinal);
        Assert.Contains("./build/verify-package-reproducibility.ps1 -Configuration Release", action, StringComparison.Ordinal);
        Assert.Contains("./build/validate-artifact-signing.ps1 -Path artifacts/publish,artifacts/packages", action, StringComparison.Ordinal);
        Assert.DoesNotContain("OnlyExo365.msi", action, StringComparison.Ordinal);
    }

    [Fact]
    public void TenantValidationWorkflow_UsesControlledAppCertificateValidationAndPublishesReport()
    {
        var workflowPath = GetRepositoryFilePath(".github", "workflows", "release-tenant-validation.yml");
        var workflow = File.ReadAllText(workflowPath);

        Assert.Contains("workflow_dispatch:", workflow, StringComparison.Ordinal);
        Assert.Contains("ONLYEXO365_CERTIFICATE_PFX_BASE64", workflow, StringComparison.Ordinal);
        Assert.Contains("Import-PfxCertificate", workflow, StringComparison.Ordinal);
        Assert.Contains("build/run-tenant-validation.ps1", workflow, StringComparison.Ordinal);
        Assert.Contains("tenant-validation-report", workflow, StringComparison.Ordinal);
        Assert.Contains("ONLYEXO365_AUTH_MODE: AppCertificate", workflow, StringComparison.Ordinal);
    }

    [Fact]
    public void TenantValidationScript_ProbesExchangeAndGraphReadOnlyCommands()
    {
        var scriptPath = GetRepositoryFilePath("build", "run-tenant-validation.ps1");
        var script = File.ReadAllText(scriptPath);

        Assert.Contains("PowerShellModuleBootstrapPolicy.json", script, StringComparison.Ordinal);
        Assert.Contains("Connect-ExchangeOnline", script, StringComparison.Ordinal);
        Assert.Contains("Get-OrganizationConfig", script, StringComparison.Ordinal);
        Assert.Contains("Get-AcceptedDomain", script, StringComparison.Ordinal);
        Assert.Contains("Get-EXOMailbox -ResultSize 3", script, StringComparison.Ordinal);
        Assert.Contains("Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -ResultSize 1", script, StringComparison.Ordinal);
        Assert.Contains("Get-ComplianceSearch -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("Invoke-CommandAvailabilityProbe -CommandName \"New-ComplianceSearchAction\"", script, StringComparison.Ordinal);
        Assert.Contains("Invoke-CommandAvailabilityProbe -CommandName \"New-CaseHoldPolicy\"", script, StringComparison.Ordinal);
        Assert.Contains("Invoke-CommandAvailabilityProbe -CommandName \"New-CaseHoldRule\"", script, StringComparison.Ordinal);
        Assert.Contains("Get-DkimSigningConfig -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("Get-HostedContentFilterPolicy -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("Get-AntiPhishPolicy -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("Get-MalwareFilterPolicy -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("Get-QuarantinePolicy -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("Get-HostedOutboundSpamFilterPolicy -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("Connect-MgGraph", script, StringComparison.Ordinal);
        Assert.Contains("Get-MgOrganization", script, StringComparison.Ordinal);
        Assert.Contains("Get-MgSubscribedSku", script, StringComparison.Ordinal);
        Assert.Contains("failed_probe_count", script, StringComparison.Ordinal);
        Assert.DoesNotContain("Set-Mailbox", script, StringComparison.Ordinal);
        Assert.DoesNotContain("New-Mailbox", script, StringComparison.Ordinal);
        Assert.DoesNotContain("Set-MgUserLicense", script, StringComparison.Ordinal);
    }

    [Fact]
    public void TenantValidationScript_ProjectsPolicyModulesAsPsCustomObjectsBeforeGrouping()
    {
        var scriptPath = GetRepositoryFilePath("build", "run-tenant-validation.ps1");
        var script = File.ReadAllText(scriptPath);

        Assert.Contains("$requiredModules.Add([pscustomobject]@{", script, StringComparison.Ordinal);
        Assert.Contains("Group-Object module", script, StringComparison.Ordinal);
        Assert.Contains("[pscustomobject]@{", script, StringComparison.Ordinal);
        Assert.DoesNotContain("$requiredModules.Add([ordered]@{", script, StringComparison.Ordinal);
    }

    [Fact]
    public void TenantValidationScript_FallsBackToDeviceCodeForInteractiveAuthInNonUiHosts()
    {
        var scriptPath = GetRepositoryFilePath("build", "run-tenant-validation.ps1");
        var script = File.ReadAllText(scriptPath);

        Assert.Contains("function Test-NonUiPowerShellHost", script, StringComparison.Ordinal);
        Assert.Contains("function Resolve-EffectiveAuthenticationMode", script, StringComparison.Ordinal);
        Assert.Contains("Falling back to DeviceCode", script, StringComparison.Ordinal);
        Assert.Contains("effective_authentication_mode", script, StringComparison.Ordinal);
        Assert.Contains("used_device_code_fallback", script, StringComparison.Ordinal);
    }

    private static string GetRepositoryFilePath(params string[] segments)
    {
        return TestPathHelper.GetRepositoryPath(segments);
    }
}

