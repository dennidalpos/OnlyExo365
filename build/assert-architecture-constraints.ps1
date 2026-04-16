[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot

$requiredFiles = @(
    'src/ExchangeAdmin.Presentation/Bootstrap/AppCompositionRoot.cs',
    'src/ExchangeAdmin.Presentation/Bootstrap/AppModuleCatalog.cs',
    'src/ExchangeAdmin.Presentation/Bootstrap/AppModuleFactory.cs',
    'src/ExchangeAdmin.Presentation/Bootstrap/AppShellModuleRegistrar.cs',
    'src/ExchangeAdmin.Presentation/Bootstrap/AppPageLoaderCatalog.cs',
    'src/ExchangeAdmin.Presentation/Bootstrap/AppPageLoadCoordinator.cs',
    'src/ExchangeAdmin.Presentation/Bootstrap/AppRuntimeContext.cs',
    'src/ExchangeAdmin.Contracts/Security/ProtectedSecretStore.cs',
    'src/ExchangeAdmin.Infrastructure/Ipc/PreparedIpcPayload.cs',
    'src/ExchangeAdmin.Worker/Operations/WorkerSecretResolver.cs',
    'build/create-setup-exe.ps1',
    'installer/ExchangeAdmin.iss'
)

foreach ($relativePath in $requiredFiles) {
    $absolutePath = Join-Path $repoRoot $relativePath
    if (-not (Test-Path -LiteralPath $absolutePath)) {
        throw "Missing required architecture file: $relativePath"
    }
}

$lineCountConstraints = @(
    @{ Path = 'src/ExchangeAdmin.Presentation/App.xaml.cs'; MaxLines = 80; Message = 'Keep WPF startup/shutdown host thin and move wiring into Bootstrap/.' },
    @{ Path = 'src/ExchangeAdmin.Presentation/Bootstrap/AppCompositionRoot.cs'; MaxLines = 80; Message = 'Keep AppCompositionRoot as a facade; move module graph and registration details into focused bootstrap helpers.' },
    @{ Path = 'src/ExchangeAdmin.Presentation/Bootstrap/AppModuleFactory.cs'; MaxLines = 120; Message = 'Keep the bootstrap factory focused on object graph creation only.' },
    @{ Path = 'src/ExchangeAdmin.Presentation/Bootstrap/AppShellModuleRegistrar.cs'; MaxLines = 190; Message = 'Keep shell registration split by responsibility; do not collapse bootstrap wiring back into a monolith.' },
    @{ Path = 'src/ExchangeAdmin.Presentation/Bootstrap/AppPageLoaderCatalog.cs'; MaxLines = 80; Message = 'Keep page loader wiring isolated and compact.' }
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
