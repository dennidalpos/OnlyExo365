#Requires -Version 7.0

  
         
                                            

            
                                                                                                
                                                                          

                        
                                                              

                
                                                                             

                  
                                           

                        
                                                             
                                                                                

              
                                                                        

                  
                                              

        
                                               
                                         

        
                                              
                                             

        
                                             
                                     

        
                             
                                     

      
                  
                
                   
                                   
                                 

                
                                        
                                                                               
  

[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    [switch]$Clean = $true,

    [switch]$Publish = $true,

    [switch]$SelfContained = $true,

    [ValidateNotNullOrEmpty()]
    [string]$RuntimeIdentifier = 'win-x64',

    [switch]$Msi = $false,

    [string]$ExportDirPath,

    [string]$ImportDirPath
)

                                   
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

       
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$SolutionDir = Split-Path -Parent $ScriptDir
$SolutionFile = Join-Path $SolutionDir "ExchangeAdmin.sln"
$OutputDir = Join-Path $SolutionDir "artifacts"
$PublishDir = Join-Path $OutputDir "publish"
$ExportDir = if ([string]::IsNullOrWhiteSpace($ExportDirPath)) { Join-Path $OutputDir "exports" } else { $ExportDirPath }
$ImportDir = if ([string]::IsNullOrWhiteSpace($ImportDirPath)) { Join-Path $OutputDir "imports" } else { $ImportDirPath }
$InstallerDir = Join-Path $OutputDir "installer"
$InstallerSourceDir = Join-Path $SolutionDir "installer"
$WixBin = "C:\Program Files (x86)\WiX Toolset v3.14\bin"

                       
$BuildStartTime = Get-Date

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host ">> $Message" -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)
    Write-Host "   [OK] $Message" -ForegroundColor Green
}

function Write-Info {
    param([string]$Message)
    Write-Host "   $Message" -ForegroundColor Gray
}

function Write-Warn {
    param([string]$Message)
    Write-Host "   [WARN] $Message" -ForegroundColor Yellow
}

function Stop-WithError {
    param([string]$Message, [int]$ExitCode = 1)
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "[FAILED] $Message" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    exit $ExitCode
}

function Get-MsiVersion {
    param([string]$VersionString)

    if ([string]::IsNullOrWhiteSpace($VersionString)) {
        return "1.0.0"
    }

    $cleanVersion = $VersionString.Split('+')[0].Split(' ')[0]

    try {
        $parsed = [Version]$cleanVersion
    }
    catch {
        return "1.0.0"
    }

    $build = if ($parsed.Build -ge 0) { $parsed.Build } else { 0 }
    return "$($parsed.Major).$($parsed.Minor).$build"
}

function Invoke-DotNet {
    param(
        [string]$Command,
        [string[]]$Arguments,
        [string]$ErrorMessage
    )

    $allArgs = @($Command) + $Arguments

    if ($VerbosePreference -eq 'Continue') {
        Write-Info "dotnet $($allArgs -join ' ')"
        & dotnet @allArgs
    }
    else {
        & dotnet @allArgs 2>&1 | ForEach-Object {
            if ($_ -match 'error') {
                Write-Host "   $_" -ForegroundColor Red
            }
            elseif ($_ -match 'warning') {
                Write-Host "   $_" -ForegroundColor Yellow
            }
            elseif ($VerbosePreference -eq 'Continue') {
                Write-Host "   $_" -ForegroundColor Gray
            }
        }
    }

    if ($LASTEXITCODE -ne 0) {
        Stop-WithError "$ErrorMessage (exit code: $LASTEXITCODE)" $LASTEXITCODE
    }
}

function Test-DotNetAvailable {
    try {
        $null = & dotnet --version 2>&1
        return $LASTEXITCODE -eq 0
    }
    catch {
        return $false
    }
}

        
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " ExchangeAdmin Build Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Configuration : $Configuration"
Write-Host "Solution      : $SolutionFile"
Write-Host "Output        : $OutputDir"
Write-Host "Exports       : $ExportDir"
if (-not [System.IO.Path]::IsPathRooted($ExportDir)) {
    $ExportDir = Join-Path $SolutionDir $ExportDir
    Write-Info "Resolved export path to absolute: $ExportDir"
}
Write-Host "Imports       : $ImportDir"
if (-not [System.IO.Path]::IsPathRooted($ImportDir)) {
    $ImportDir = Join-Path $SolutionDir $ImportDir
    Write-Info "Resolved import path to absolute: $ImportDir"
}
Write-Host "Publish       : $($Publish.IsPresent)"
Write-Host "Self-contained: $($SelfContained.IsPresent)"
Write-Host "Runtime       : $RuntimeIdentifier"
Write-Host "MSI           : $($Msi.IsPresent)"
Write-Host "Started at    : $($BuildStartTime.ToString('HH:mm:ss'))"

                        
if (-not (Test-Path $SolutionFile)) {
    Stop-WithError "Solution file not found: $SolutionFile"
}

                 
Write-Step "Checking prerequisites"

try {
    if (-not (Test-DotNetAvailable)) {
        throw "dotnet not found"
    }
    $dotnetVersion = & dotnet --version 2>&1
    Write-Success ".NET SDK version: $dotnetVersion"

                         
    $majorVersion = [int]($dotnetVersion.Split('.')[0])
    if ($majorVersion -lt 8) {
        Stop-WithError ".NET 8 SDK or later is required. Found: $dotnetVersion"
    }
}
catch {
    Stop-WithError ".NET SDK not found. Please install .NET 8 SDK from https://dotnet.microsoft.com/download"
}

                                           
if ($PSVersionTable.Platform -and $PSVersionTable.Platform -ne 'Win32NT') {
    Stop-WithError "This project requires Windows (WPF is Windows-only)"
}

       
if ($Clean) {
    Write-Step "Cleaning"

    if (Test-Path $OutputDir) {
        Remove-Item -Path $OutputDir -Recurse -Force -ErrorAction Stop
        Write-Success "Removed: $OutputDir"
    }
    else {
        Write-Info "Output directory doesn't exist, skipping"
    }

    Invoke-DotNet -Command "clean" -Arguments @($SolutionFile, "-c", $Configuration, "--verbosity", "minimal") `
                  -ErrorMessage "Clean failed"
    Write-Success "Solution cleaned"
}

         
Write-Step "Restoring NuGet packages"

$restoreArgs = @($SolutionFile, "--verbosity", "minimal")

                                                          
if ($SelfContained) {
    $restoreArgs += '-r', $RuntimeIdentifier
    Write-Info "Restoring for runtime: $RuntimeIdentifier"
}

Invoke-DotNet -Command "restore" -Arguments $restoreArgs `
              -ErrorMessage "Package restore failed"
Write-Success "Packages restored"

       
Write-Step "Building solution"

$buildArgs = @(
    $SolutionFile
    "-c", $Configuration
    "--no-restore"
    "-warnaserror:nullable"
    "-p:TreatWarningsAsErrors=false"
)

if ($VerbosePreference -ne 'Continue') {
    $buildArgs += "--verbosity", "minimal"
}

Invoke-DotNet -Command "build" -Arguments $buildArgs -ErrorMessage "Build failed"
Write-Success "Build succeeded"

         
if ($Publish) {
    Write-Step "Publishing applications"

                              
    if (-not (Test-Path $PublishDir)) {
        New-Item -Path $PublishDir -ItemType Directory -Force | Out-Null
        Write-Info "Created: $PublishDir"
    }

    foreach ($pathInfo in @(@{ Path = $ExportDir; Label = "exports" }, @{ Path = $ImportDir; Label = "imports" })) {
        if (-not (Test-Path $pathInfo.Path)) {
            New-Item -Path $pathInfo.Path -ItemType Directory -Force | Out-Null
            Write-Info "Created $($pathInfo.Label) directory: $($pathInfo.Path)"
        }
    }

    $publishArgs = @(
        "-c", $Configuration
        "-o", $PublishDir
    )

    if ($SelfContained) {
                                                                              
        $publishArgs += '--self-contained', 'true'
        $publishArgs += '-r', $RuntimeIdentifier
        $publishArgs += '-p:PublishSingleFile=false'
        Write-Info "Mode: Self-contained ($RuntimeIdentifier)"
    }
    else {
                                                              
        $publishArgs += "--no-build"
        $publishArgs += "--self-contained", "false"
        Write-Info "Mode: Framework-dependent"
    }

    if ($VerbosePreference -ne 'Continue') {
        $publishArgs += "--verbosity", "minimal"
    }

                              
    Write-Info "Publishing ExchangeAdmin.Presentation..."
    $presentationProject = Join-Path $SolutionDir "src\ExchangeAdmin.Presentation\ExchangeAdmin.Presentation.csproj"

    if (-not (Test-Path $presentationProject)) {
        Stop-WithError "Presentation project not found: $presentationProject"
    }

    Invoke-DotNet -Command "publish" -Arguments (@($presentationProject) + $publishArgs) `
                  -ErrorMessage "Publish of Presentation failed"
    Write-Success "ExchangeAdmin.Presentation published"

                    
    Write-Info "Publishing ExchangeAdmin.Worker..."
    $workerProject = Join-Path $SolutionDir "src\ExchangeAdmin.Worker\ExchangeAdmin.Worker.csproj"

    if (-not (Test-Path $workerProject)) {
        Stop-WithError "Worker project not found: $workerProject"
    }

    Invoke-DotNet -Command "publish" -Arguments (@($workerProject) + $publishArgs) `
                  -ErrorMessage "Publish of Worker failed"
    Write-Success "ExchangeAdmin.Worker published"

                          
    $publishedFiles = Get-ChildItem -Path $PublishDir -Filter "*.exe" | Select-Object -ExpandProperty Name
    Write-Info "Published executables: $($publishedFiles -join ', ')"

                    
    $publishSize = (Get-ChildItem -Path $PublishDir -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB
    Write-Info "Total size: $([math]::Round($publishSize, 2)) MB"
}

if ($Msi) {
    Write-Step "Building MSI installer"

    if (-not $Publish) {
        Stop-WithError "MSI build requires published output. Run with -Publish."
    }

    if (-not (Test-Path $PublishDir)) {
        Stop-WithError "Publish output not found: $PublishDir"
    }

    $presentationExe = Join-Path $PublishDir "ExchangeAdmin.Presentation.exe"
    if (-not (Test-Path $presentationExe)) {
        Stop-WithError "Presentation executable not found: $presentationExe"
    }

    $wixCandle = Join-Path $WixBin "candle.exe"
    $wixLight = Join-Path $WixBin "light.exe"
    $wixHeat = Join-Path $WixBin "heat.exe"

    foreach ($tool in @($wixCandle, $wixLight, $wixHeat)) {
        if (-not (Test-Path $tool)) {
            Stop-WithError "WiX tool not found: $tool"
        }
    }

    if (-not (Test-Path $InstallerSourceDir)) {
        Stop-WithError "Installer source directory not found: $InstallerSourceDir"
    }

    $wixSource = Join-Path $InstallerSourceDir "ExchangeAdmin.wxs"
    if (-not (Test-Path $wixSource)) {
        Stop-WithError "WiX source file not found: $wixSource"
    }

    if (-not (Test-Path $InstallerDir)) {
        New-Item -Path $InstallerDir -ItemType Directory -Force | Out-Null
        Write-Info "Created: $InstallerDir"
    }

    $wixObjDir = Join-Path $InstallerDir "obj"
    if (-not (Test-Path $wixObjDir)) {
        New-Item -Path $wixObjDir -ItemType Directory -Force | Out-Null
    }

    $harvestFile = Join-Path $InstallerDir "PublishFiles.wxs"
    $msiOutput = Join-Path $InstallerDir "ExchangeAdmin.msi"

    $rawVersion = (Get-Item $presentationExe).VersionInfo.ProductVersion
    $msiVersion = Get-MsiVersion -VersionString $rawVersion
    Write-Info "MSI version: $msiVersion"

    Write-Info "Harvesting publish directory..."
    & $wixHeat "dir" $PublishDir `
        "-cg" "PublishFiles" `
        "-dr" "INSTALLDIR" `
        "-gg" `
        "-srd" `
        "-sfrag" `
        "-sreg" `
        "-var" "var.PublishDir" `
        "-out" $harvestFile

    if ($LASTEXITCODE -ne 0) {
        Stop-WithError "Heat harvesting failed (exit code: $LASTEXITCODE)" $LASTEXITCODE
    }

    Write-Info "Compiling WiX sources..."
    & $wixCandle "-nologo" `
        "-arch" "x64" `
        "-dPublishDir=$PublishDir" `
        "-dProductName=ExchangeAdmin" `
        "-dManufacturer=OnlyExo365" `
        "-dProductVersion=$msiVersion" `
        "-out" "$wixObjDir\" `
        $wixSource `
        $harvestFile

    if ($LASTEXITCODE -ne 0) {
        Stop-WithError "WiX candle failed (exit code: $LASTEXITCODE)" $LASTEXITCODE
    }

    $mainObj = Join-Path $wixObjDir "ExchangeAdmin.wixobj"
    $harvestObj = Join-Path $wixObjDir "PublishFiles.wixobj"

    Write-Info "Linking MSI..."
    & $wixLight "-nologo" `
        "-ext" "WixUIExtension" `
        "-out" $msiOutput `
        $mainObj `
        $harvestObj

    if ($LASTEXITCODE -ne 0) {
        Stop-WithError "WiX light failed (exit code: $LASTEXITCODE)" $LASTEXITCODE
    }

    Write-Success "MSI created: $msiOutput"
}

         
$BuildEndTime = Get-Date
$BuildDuration = $BuildEndTime - $BuildStartTime

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " BUILD COMPLETED SUCCESSFULLY" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Duration: $($BuildDuration.TotalSeconds.ToString('F1')) seconds" -ForegroundColor Gray

if ($Publish) {
    Write-Host ""
    Write-Host "Published to: $PublishDir" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "To run the application:" -ForegroundColor Yellow
    Write-Host "  cd `"$PublishDir`""
    Write-Host "  .\ExchangeAdmin.Presentation.exe"
    Write-Host ""
    Write-Host "Runtime prerequisites:" -ForegroundColor Yellow
    Write-Host "  1. PowerShell 7+ (pwsh.exe in PATH)"
    Write-Host "  2. ExchangeOnlineManagement module:"
    Write-Host "     Install-Module ExchangeOnlineManagement -Scope CurrentUser"

    if (-not $SelfContained) {
        Write-Host "  3. .NET 8 Runtime (framework-dependent build)"
    }
}

if ($Msi) {
    Write-Host ""
    Write-Host "MSI output: $InstallerDir" -ForegroundColor Cyan
}

Write-Host ""
exit 0
