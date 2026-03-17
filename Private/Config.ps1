<#
.SYNOPSIS
    Configuration management for the Azure DevOps Time Tracker.
.DESCRIPTION
    Stores and loads configuration (organization, project, PAT) from a JSON file
    in the user's config directory.
#>

function Get-TTConfigDir {
    if ($env:AZDOTT_CONFIG_DIR) {
        return $env:AZDOTT_CONFIG_DIR
    }
    if ($env:XDG_CONFIG_HOME) {
        return Join-Path $env:XDG_CONFIG_HOME "AzDoTimeTracker"
    }
    if ($IsWindows) {
        return Join-Path $env:APPDATA "AzDoTimeTracker"
    }
    return Join-Path $HOME ".config" "AzDoTimeTracker"
}

function Get-TTConfigPath {
    return Join-Path (Get-TTConfigDir) "config.json"
}

function Get-TTConfig {
    $configPath = Get-TTConfigPath
    if (Test-Path $configPath) {
        $config = Get-Content $configPath -Raw | ConvertFrom-Json
        return $config
    }
    return $null
}

function Save-TTConfig {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        [int]$PriParentId = 0,
        [string]$PriParentTitle = "",
        [string]$ScrumLanguage = 'en'
    )

    $configDir = Get-TTConfigDir
    if (-not (Test-Path $configDir)) {
        New-Item -ItemType Directory -Path $configDir -Force | Out-Null
    }

    $config = @{
        Organization   = $Organization
        Project        = $Project
        PAT            = $PAT
        PriParentId    = $PriParentId
        PriParentTitle = $PriParentTitle
        ScrumLanguage  = $ScrumLanguage
    }
    $configPath = Get-TTConfigPath
    $config | ConvertTo-Json | Set-Content $configPath
    return $config
}

function Save-PriParentConfig {
    param(
        [int]$ParentId,
        [string]$ParentTitle
    )
    $existing = Get-TTConfig
    $org   = if ($existing -and $existing.Organization) { $existing.Organization } else { "" }
    $proj  = if ($existing -and $existing.Project)      { $existing.Project      } else { "" }
    $pat   = if ($existing -and $existing.PAT)          { $existing.PAT          } else { "" }
    $lang  = if ($existing -and $existing.ScrumLanguage) { $existing.ScrumLanguage } else { 'en' }
    Save-TTConfig -Organization $org -Project $proj -PAT $pat `
        -PriParentId $ParentId -PriParentTitle $ParentTitle -ScrumLanguage $lang | Out-Null
}

function Save-ScrumLanguageConfig {
    param([string]$Language = 'en')
    $existing = Get-TTConfig
    $org   = if ($existing -and $existing.Organization)  { $existing.Organization  } else { "" }
    $proj  = if ($existing -and $existing.Project)       { $existing.Project       } else { "" }
    $pat   = if ($existing -and $existing.PAT)           { $existing.PAT           } else { "" }
    $pid2  = if ($existing -and $existing.PriParentId)   { [int]$existing.PriParentId } else { 0 }
    $ptit  = if ($existing -and $existing.PriParentTitle){ [string]$existing.PriParentTitle } else { "" }
    Save-TTConfig -Organization $org -Project $proj -PAT $pat `
        -PriParentId $pid2 -PriParentTitle $ptit -ScrumLanguage $Language | Out-Null
}

function Initialize-TTConfig {
    $config = Get-TTConfig
    if ($null -ne $config -and $config.Organization -and $config.Project -and $config.PAT) {
        return $config
    }

    return Request-TTConfig
}

function Request-TTConfig {
    $existing = Get-TTConfig

    Write-Host ""
    Write-Host "=== Azure DevOps Time Tracker Setup ===" -ForegroundColor Cyan
    Write-Host ""

    $defaultOrg = if ($existing -and $existing.Organization) { $existing.Organization } else { "" }
    $defaultProj = if ($existing -and $existing.Project) { $existing.Project } else { "" }
    $defaultPat = if ($existing -and $existing.PAT) { $existing.PAT } else { "" }

    if ($defaultOrg) {
        $org = Read-Host "Organization [$defaultOrg]"
        if (-not $org) { $org = $defaultOrg }
    } else {
        $org = Read-Host "Enter your Azure DevOps Organization name (e.g. 'myorg')"
    }

    if ($defaultProj) {
        $proj = Read-Host "Project [$defaultProj]"
        if (-not $proj) { $proj = $defaultProj }
    } else {
        $proj = Read-Host "Enter your Azure DevOps Project name"
    }

    if ($defaultPat) {
        $maskedPat = $defaultPat.Substring(0, [Math]::Min(4, $defaultPat.Length)) + "****"
        $pat = Read-Host "PAT [$maskedPat]" -MaskInput
        if (-not $pat) { $pat = $defaultPat }
    } else {
        $pat = Read-Host "Enter your Personal Access Token (PAT)" -MaskInput
    }

    if (-not $org -or -not $proj -or -not $pat) {
        Write-Host "All fields are required. Exiting." -ForegroundColor Red
        exit 1
    }

    $existingParentId    = if ($existing -and $existing.PriParentId)    { [int]$existing.PriParentId    } else { 0  }
    $existingParentTitle = if ($existing -and $existing.PriParentTitle) { [string]$existing.PriParentTitle } else { "" }
    $existingScrumLang   = if ($existing -and $existing.ScrumLanguage)  { [string]$existing.ScrumLanguage } else { 'en' }

    $config = Save-TTConfig -Organization $org -Project $proj -PAT $pat `
        -PriParentId $existingParentId -PriParentTitle $existingParentTitle `
        -ScrumLanguage $existingScrumLang
    Write-Host "Configuration saved to: $(Get-TTConfigPath)" -ForegroundColor Green
    Write-Host ""
    return $config
}
