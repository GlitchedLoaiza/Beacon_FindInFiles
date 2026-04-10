<#
.SYNOPSIS
    Builds Beacon application as a single executable file.

.DESCRIPTION
    This script publishes the AutoPilot Log Search WPF application as a 
    self-contained single executable file for Windows x64.

.PARAMETER Configuration
    Build configuration (Debug or Release). Default is Release.

.PARAMETER Platform
    Target platform (win-x64, win-x86, win-arm64). Default is win-x64.

.PARAMETER OutputFolder
    Optional custom output folder for the final executable.

.PARAMETER Clean
    Clean previous build artifacts before building.

.EXAMPLE
    .\Build-SingleExe.ps1
    
.EXAMPLE
    .\Build-SingleExe.ps1 -Configuration Debug -Clean
    
.EXAMPLE
    .\Build-SingleExe.ps1 -Platform win-x86 -OutputFolder "C:\Deploy"
#>

param(
    [Parameter()]
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    
    [Parameter()]
    [ValidateSet("win-x64", "win-x86", "win-arm64")]
    [string]$Platform = "win-x64",
    
    [Parameter()]
    [string]$OutputFolder = "",
    
    [Parameter()]
    [switch]$Clean
)

# Script configuration
$ErrorActionPreference = "Stop"
$projectPath = "AutoPilot Log Search\AutoPilot Log Search.vbproj"
$exeName = "Beacon.exe"

# Display header
Write-Host "`n================================================" -ForegroundColor Cyan
Write-Host "  Beacon Single-File EXE Builder" -ForegroundColor Cyan
Write-Host "================================================`n" -ForegroundColor Cyan

# Verify project file exists
if (-not (Test-Path $projectPath)) {
    Write-Host "ERROR: Project file not found: $projectPath" -ForegroundColor Red
    Write-Host "Please run this script from the solution root directory." -ForegroundColor Yellow
    exit 1
}

Write-Host "Configuration: $Configuration" -ForegroundColor Green
Write-Host "Platform:      $Platform" -ForegroundColor Green
Write-Host "Project:       $projectPath" -ForegroundColor Green

# Clean if requested
if ($Clean) {
    Write-Host "`nCleaning previous builds..." -ForegroundColor Yellow
    
    $binPath = "AutoPilot Log Search\bin"
    $objPath = "AutoPilot Log Search\obj"
    
    if (Test-Path $binPath) {
        Remove-Item $binPath -Recurse -Force
        Write-Host "  ✓ Removed bin folder" -ForegroundColor Gray
    }
    
    if (Test-Path $objPath) {
        Remove-Item $objPath -Recurse -Force
        Write-Host "  ✓ Removed obj folder" -ForegroundColor Gray
    }
}

# Build the project
Write-Host "`nBuilding project..." -ForegroundColor Yellow
try {
    dotnet build $projectPath -c $Configuration --nologo
    if ($LASTEXITCODE -ne 0) {
        throw "Build failed with exit code $LASTEXITCODE"
    }
    Write-Host "  ✓ Build successful" -ForegroundColor Green
} catch {
    Write-Host "  ✗ Build failed: $_" -ForegroundColor Red
    exit 1
}

# Publish as single file
Write-Host "`nPublishing single-file executable..." -ForegroundColor Yellow
$publishPath = "AutoPilot Log Search\bin\$Configuration\net10.0-windows\$Platform\publish"

try {
    dotnet publish $projectPath `
        -c $Configuration `
        -r $Platform `
        --self-contained `
        -p:PublishSingleFile=true `
        -p:IncludeNativeLibrariesForSelfExtract=true `
        -p:PublishReadyToRun=true `
        --nologo
        
    if ($LASTEXITCODE -ne 0) {
        throw "Publish failed with exit code $LASTEXITCODE"
    }
    Write-Host "  ✓ Publish successful" -ForegroundColor Green
} catch {
    Write-Host "  ✗ Publish failed: $_" -ForegroundColor Red
    exit 1
}

# Verify output
$publishedExe = Join-Path $publishPath $exeName
if (-not (Test-Path $publishedExe)) {
    Write-Host "`nERROR: Expected executable not found: $publishedExe" -ForegroundColor Red
    exit 1
}

$fileInfo = Get-Item $publishedExe
$fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2)

Write-Host "`n================================================" -ForegroundColor Green
Write-Host "  BUILD SUCCESSFUL!" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green
Write-Host "`nExecutable Details:" -ForegroundColor Cyan
Write-Host "  Name:     $($fileInfo.Name)" -ForegroundColor White
Write-Host "  Size:     $fileSizeMB MB" -ForegroundColor White
Write-Host "  Location: $($fileInfo.FullName)" -ForegroundColor White
Write-Host "  Modified: $($fileInfo.LastWriteTime)" -ForegroundColor White

# Copy to custom output folder if specified
if ($OutputFolder) {
    Write-Host "`nCopying to custom output folder..." -ForegroundColor Yellow
    
    if (-not (Test-Path $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
        Write-Host "  ✓ Created output folder: $OutputFolder" -ForegroundColor Gray
    }
    
    $destPath = Join-Path $OutputFolder $exeName
    Copy-Item $publishedExe $destPath -Force
    Write-Host "  ✓ Copied to: $destPath" -ForegroundColor Green
}

# Option to open folder
Write-Host "`n================================================`n" -ForegroundColor Cyan
$openFolder = Read-Host "Open output folder? (Y/N)"
if ($openFolder -eq 'Y' -or $openFolder -eq 'y') {
    if ($OutputFolder -and (Test-Path $OutputFolder)) {
        explorer.exe $OutputFolder
    } else {
        explorer.exe (Split-Path $publishedExe -Parent)
    }
}

Write-Host "`nDone!`n" -ForegroundColor Green
