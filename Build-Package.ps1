#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Builds a NuGet package (.nupkg) for the AzDoTimeTracker PowerShell module.

.DESCRIPTION
    Creates a .nupkg file identical to what Publish-Module would upload to
    the PowerShell Gallery.  The resulting package can be shared and
    installed offline with Install-Module.

    The package version is read automatically from AzDoTimeTracker.psd1.

    Requirements:
      - PowerShell 7+

.PARAMETER OutputDir
    Directory where the .nupkg file will be written. Defaults to ./out.

.EXAMPLE
    ./Build-Package.ps1

.EXAMPLE
    ./Build-Package.ps1 -OutputDir ~/Desktop
#>
[CmdletBinding()]
param(
    [string]$OutputDir = (Join-Path $PSScriptRoot 'out')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Paths ────────────────────────────────────────────────────────────
$moduleRoot   = $PSScriptRoot
$manifestPath = Join-Path $moduleRoot 'AzDoTimeTracker.psd1'

if (-not (Test-Path $manifestPath)) {
    Write-Error "Module manifest not found at $manifestPath"
    return
}

# ── Read version from manifest ───────────────────────────────────────
$manifest = Import-PowerShellDataFile $manifestPath
$version  = $manifest.ModuleVersion
Write-Host "Building AzDoTimeTracker v$version" -ForegroundColor Cyan

# ── Create a clean staging copy (exclude dev files) ──────────────────
$staging = Join-Path ([System.IO.Path]::GetTempPath()) "AzDoTimeTracker-nupkg-$([guid]::NewGuid().ToString('N').Substring(0,8))"
$stageModule = Join-Path $staging 'AzDoTimeTracker'
New-Item -ItemType Directory -Path $stageModule -Force | Out-Null

Write-Host "Staging module files..." -ForegroundColor Gray

# Files to include in the package
$includeFiles = @(
    'AzDoTimeTracker.psd1'
    'AzDoTimeTracker.psm1'
    'README.md'
)

foreach ($f in $includeFiles) {
    $src = Join-Path $moduleRoot $f
    if (Test-Path $src) {
        Copy-Item $src -Destination $stageModule
    }
}

foreach ($dir in @('Public', 'Private')) {
    $src = Join-Path $moduleRoot $dir
    if (Test-Path $src) {
        Copy-Item $src -Destination (Join-Path $stageModule $dir) -Recurse
    }
}

# ── Create a local file-based NuGet repository ──────────────────────
$localRepo = Join-Path $staging 'repo'
New-Item -ItemType Directory -Path $localRepo -Force | Out-Null

$repoName = "AzDoTT-Build-$([guid]::NewGuid().ToString('N').Substring(0,8))"

try {
    Register-PSRepository -Name $repoName -SourceLocation $localRepo `
        -PublishLocation $localRepo -InstallationPolicy Trusted

    Write-Host "Creating .nupkg via Publish-Module..." -ForegroundColor Gray

    Publish-Module -Path $stageModule -Repository $repoName -Force

    # ── Copy .nupkg to output ────────────────────────────────────────
    if (-not (Test-Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }

    $nupkg = Get-ChildItem -Path $localRepo -Filter '*.nupkg' | Select-Object -First 1
    if (-not $nupkg) {
        Write-Error "Publish-Module succeeded but no .nupkg was found in $localRepo"
        return
    }

    $destPath = Join-Path $OutputDir $nupkg.Name
    Copy-Item $nupkg.FullName -Destination $destPath -Force

    $sizeKB = [Math]::Round($nupkg.Length / 1024, 1)
    Write-Host ""
    Write-Host "Package created: $destPath ($sizeKB KB)" -ForegroundColor Green
    Write-Host ""
    Write-Host "To install locally:" -ForegroundColor Yellow
    Write-Host "  # Register a local repo pointing to the folder containing the .nupkg" -ForegroundColor Gray
    Write-Host "  Register-PSRepository -Name Local -SourceLocation '$((Resolve-Path $OutputDir).Path)' -InstallationPolicy Trusted" -ForegroundColor White
    Write-Host "  Install-Module -Name AzDoTimeTracker -Repository Local" -ForegroundColor White
    Write-Host ""
    Write-Host "Or copy directly to your modules path:" -ForegroundColor Yellow
    Write-Host "  `$dest = Join-Path (`$env:PSModulePath -split '[;:]')[0] 'AzDoTimeTracker'" -ForegroundColor White
    Write-Host "  Copy-Item -Path '$stageModule' -Destination `$dest -Recurse" -ForegroundColor White
}
finally {
    # ── Clean up ─────────────────────────────────────────────────────
    Unregister-PSRepository -Name $repoName -ErrorAction SilentlyContinue
    Remove-Item $staging -Recurse -Force -ErrorAction SilentlyContinue
}
