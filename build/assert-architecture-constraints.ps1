[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot

$requiredFiles = @(
    'src/OnlyExo365.Shell/Bootstrap/AppCompositionRoot.cs',
    'src/OnlyExo365.Shell/Bootstrap/AppModuleCatalog.cs',
    'src/OnlyExo365.Shell/Bootstrap/AppModuleFactory.cs',
    'src/OnlyExo365.Shell/Bootstrap/AppShellModuleRegistrar.cs',
    'src/OnlyExo365.Shell/Bootstrap/AppPageLoaderCatalog.cs',
    'src/OnlyExo365.Shell/Bootstrap/AppPageLoadCoordinator.cs',
    'src/OnlyExo365.Shell/Bootstrap/AppRuntimeContext.cs',
    'src/OnlyExo365.Contracts/Security/ProtectedSecretStore.cs',
    'src/OnlyExo365.Shell/Ipc/PreparedIpcPayload.cs',
    'src/OnlyExo365.Worker/Operations/WorkerSecretResolver.cs',
    'build/create-setup-exe.ps1',
    'installer/OnlyExo365.iss'
)

$requiredCharacterizationTests = @(
    'tests/OnlyExo365.Tests/IpcContractsTests.cs',
    'tests/OnlyExo365.Tests/IpcSecretHandlingTests.cs',
    'tests/OnlyExo365.Tests/IpcServerTests.cs',
    'tests/OnlyExo365.Tests/IpcServerOversizeTests.cs',
    'tests/OnlyExo365.Tests/OperationDispatcherCharacterizationTests.cs',
    'tests/OnlyExo365.Tests/OperationDispatcherMessagingTests.cs',
    'tests/OnlyExo365.Tests/PersistentLogWriterTests.cs',
    'tests/OnlyExo365.Tests/WorkerSupervisorTests.cs'
)

foreach ($relativePath in $requiredFiles) {
    $absolutePath = Join-Path $repoRoot $relativePath
    if (-not (Test-Path -LiteralPath $absolutePath)) {
        throw "Missing required architecture file: $relativePath"
    }
}

foreach ($relativePath in $requiredCharacterizationTests) {
    $absolutePath = Join-Path $repoRoot $relativePath
    if (-not (Test-Path -LiteralPath $absolutePath)) {
        throw "Missing required characterization test for IPC/orchestration boundary: $relativePath"
    }
}

$lineCountConstraints = @(
    @{ Path = 'src/OnlyExo365.Shell/App.xaml.cs'; MaxLines = 80; Message = 'Keep WPF startup/shutdown host thin and move wiring into Bootstrap/.' },
    @{ Path = 'src/OnlyExo365.Shell/Bootstrap/AppCompositionRoot.cs'; MaxLines = 80; Message = 'Keep AppCompositionRoot as a facade; move module graph and registration details into focused bootstrap helpers.' },
    @{ Path = 'src/OnlyExo365.Shell/Bootstrap/AppModuleFactory.cs'; MaxLines = 120; Message = 'Keep the bootstrap factory focused on object graph creation only.' },
    @{ Path = 'src/OnlyExo365.Shell/Bootstrap/AppShellModuleRegistrar.cs'; MaxLines = 190; Message = 'Keep shell registration split by responsibility; do not collapse bootstrap wiring back into a monolith.' },
    @{ Path = 'src/OnlyExo365.Shell/Bootstrap/AppPageLoaderCatalog.cs'; MaxLines = 80; Message = 'Keep page loader wiring isolated and compact.' },
    @{ Path = 'src/OnlyExo365.Shell/Services/WorkerService.cs'; MaxLines = 720; Message = 'WorkerService is an orchestration hotspot; split by worker facade responsibilities before adding more behavior.' },
    @{ Path = 'src/OnlyExo365.Shell/Ipc/WorkerClient.cs'; MaxLines = 700; Message = 'WorkerClient is an IPC hotspot; prefer facet clients over expanding the runtime client.' },
    @{ Path = 'src/OnlyExo365.Shell/Ipc/WorkerSupervisor.cs'; MaxLines = 720; Message = 'WorkerSupervisor is a process-lifecycle hotspot; separate lifecycle and console/log replay responsibilities before growth.' },
    @{ Path = 'src/OnlyExo365.Worker/Ipc/IpcServer.cs'; MaxLines = 760; Message = 'IpcServer is a transport/dispatch hotspot; keep transport and request dispatch from growing together.' },
    @{ Path = 'src/OnlyExo365.Worker/Operations/OperationDispatcher.cs'; MaxLines = 110; Message = 'Keep OperationDispatcher core as routing only; add behavior in focused handlers.' },
    @{ Path = 'src/OnlyExo365.Worker/Operations/OperationDispatcher.Mailboxes.cs'; MaxLines = 430; Message = 'Mailbox operation dispatch is a hotspot; split by mailbox facet before expanding.' },
    @{ Path = 'src/OnlyExo365.Worker/Operations/OperationDispatcher.MailFlow.cs'; MaxLines = 370; Message = 'Mail-flow operation dispatch is a hotspot; split by mail-flow facet before expanding.' },
    @{ Path = 'src/OnlyExo365.Worker/Operations/OperationDispatcher.Recipients.cs'; MaxLines = 580; Message = 'Recipient operation dispatch is a hotspot; split by recipient facet before expanding.' },
    @{ Path = 'src/OnlyExo365.Worker/PowerShell/ExoMailboxScriptFactory.cs'; MaxLines = 840; Message = 'Mailbox script generation is a hotspot; add focused command factories instead of growing this file.' },
    @{ Path = 'src/OnlyExo365.Worker/PowerShell/ExoComplianceCommands.cs'; MaxLines = 760; Message = 'Compliance command generation is a hotspot; split by compliance facet before expanding.' },
    @{ Path = 'src/OnlyExo365.Worker/PowerShell/ExoMailboxLicenseCommands.cs'; MaxLines = 720; Message = 'License command generation is a hotspot; split by license facet before expanding.' },
    @{ Path = 'src/OnlyExo365.Worker/PowerShell/ExoMobileDeviceCommands.cs'; MaxLines = 700; Message = 'Mobile-device command generation is a hotspot; split by device facet before expanding.' },
    @{ Path = 'src/OnlyExo365.Worker/PowerShell/ExoMailboxReportingCommands.cs'; MaxLines = 690; Message = 'Mailbox reporting command generation is a hotspot; split reporting facets before expanding.' },
    @{ Path = 'src/OnlyExo365.Worker/PowerShell/ExoMigrationCommands.cs'; MaxLines = 600; Message = 'Migration command generation is a hotspot; split migration facets before expanding.' },
    @{ Path = 'src/OnlyExo365.Worker/PowerShell/ExoPermissionCommands.cs'; MaxLines = 520; Message = 'Permission command generation is a hotspot; split permission facets before expanding.' },
    @{ Path = 'src/OnlyExo365.Worker/PowerShell/ExoSupportCommands.cs'; MaxLines = 520; Message = 'Support command generation is a hotspot; split support facets before expanding.' },
    @{ Path = 'src/OnlyExo365.Worker/PowerShell/ErrorClassifier.cs'; MaxLines = 520; Message = 'Error classification is a hotspot; split classifiers by error family before expanding.' }
)

$results = foreach ($constraint in $lineCountConstraints) {
    $absolutePath = Join-Path $repoRoot $constraint.Path
    $lineCount = (Get-Content -LiteralPath $absolutePath | Measure-Object -Line).Lines
    if ($lineCount -gt $constraint.MaxLines) {
        throw "$($constraint.Path) regressed to $lineCount lines. $($constraint.Message)"
    }

    [pscustomobject]@{
        Path = $constraint.Path
        LineCount = $lineCount
        MaxLines = $constraint.MaxLines
    }
}

Write-Host "Architecture constraints passed."
foreach ($result in $results) {
    Write-Host "$($result.Path) line count: $($result.LineCount) / $($result.MaxLines)"
}

