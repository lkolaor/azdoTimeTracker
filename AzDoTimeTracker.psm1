<#
.SYNOPSIS
    Root module for AzDoTimeTracker.
.DESCRIPTION
    Dot-sources all private and public function files.
#>

$ModuleRoot = $PSScriptRoot

# Debug logging flag – set by Start-TimeTracker when -Debug is passed
$script:TTDebugEnabled = $false

# Import private (internal) functions
$privatePath = Join-Path $ModuleRoot 'Private'
if (Test-Path $privatePath) {
    $privateFiles = Get-ChildItem -Path $privatePath -Filter '*.ps1' -Recurse -ErrorAction SilentlyContinue
    foreach ($file in $privateFiles) {
        try {
            . $file.FullName
        }
        catch {
            Write-Error "Failed to import private function '$($file.FullName)': $_"
        }
    }
}

# Import public (exported) functions
$publicPath = Join-Path $ModuleRoot 'Public'
if (Test-Path $publicPath) {
    $publicFiles = Get-ChildItem -Path $publicPath -Filter '*.ps1' -Recurse -ErrorAction SilentlyContinue
    foreach ($file in $publicFiles) {
        try {
            . $file.FullName
        }
        catch {
            Write-Error "Failed to import public function '$($file.FullName)': $_"
        }
    }
}

# Export only the public functions
Export-ModuleMember -Function 'Start-TimeTracker'
