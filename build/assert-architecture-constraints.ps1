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

foreach ($relativePath in $requiredFiles) {
    $absolutePath = Join-Path $repoRoot $relativePath
    if (-not (Test-Path -LiteralPath $absolutePath)) {
        throw "Missing required architecture file: $relativePath"
    }
}

$lineCountConstraints = @(
    @{ Path = 'src/OnlyExo365.Shell/App.xaml.cs'; MaxLines = 80; Message = 'Keep WPF startup/shutdown host thin and move wiring into Bootstrap/.' },
    @{ Path = 'src/OnlyExo365.Shell/Bootstrap/AppCompositionRoot.cs'; MaxLines = 80; Message = 'Keep AppCompositionRoot as a facade; move module graph and registration details into focused bootstrap helpers.' },
    @{ Path = 'src/OnlyExo365.Shell/Bootstrap/AppModuleFactory.cs'; MaxLines = 120; Message = 'Keep the bootstrap factory focused on object graph creation only.' },
    @{ Path = 'src/OnlyExo365.Shell/Bootstrap/AppShellModuleRegistrar.cs'; MaxLines = 190; Message = 'Keep shell registration split by responsibility; do not collapse bootstrap wiring back into a monolith.' },
    @{ Path = 'src/OnlyExo365.Shell/Bootstrap/AppPageLoaderCatalog.cs'; MaxLines = 80; Message = 'Keep page loader wiring isolated and compact.' }
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

