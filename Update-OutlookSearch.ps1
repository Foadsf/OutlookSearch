<#
.SYNOPSIS
Updates the OutlookSearch module to the latest version.
#>
[CmdletBinding()]
param(
    [string]$ModulePath = "$env:USERPROFILE\Documents\PowerShell\Modules\OutlookSearch"
)

Write-Host "OutlookSearch Updater" -ForegroundColor Cyan
Write-Host "=====================`n" -ForegroundColor Cyan

# Check current version
$currentModule = Get-Module OutlookSearch -ListAvailable | Select-Object -First 1
if (-not $currentModule) {
    Write-Error "OutlookSearch not found. Please run Install-OutlookSearch.ps1 first."
    exit 1
}

Write-Host "Current version: $($currentModule.Version)" -ForegroundColor Yellow

# Get latest version info
try {
    $latestInfo = Invoke-RestMethod -Uri "https://api.github.com/repos/Foadsf/OutlookSearch/releases/latest"
    $latestVersion = [version]($latestInfo.tag_name -replace '^v', '')
    
    Write-Host "Latest version: $latestVersion" -ForegroundColor Yellow
    
    if ($latestVersion -le $currentModule.Version) {
        Write-Host "`n✓ You have the latest version!" -ForegroundColor Green
        exit 0
    }
    
    Write-Host "`nUpdate available!" -ForegroundColor Green
    $confirm = Read-Host "Proceed with update? (Y/N)"
    
    if ($confirm -eq 'Y') {
        # Backup current
        $backupPath = "$ModulePath.backup.$(Get-Date -Format 'yyyyMMddHHmmss')"
        Copy-Item -Path $ModulePath -Destination $backupPath -Recurse
        
        # Remove current
        Remove-Item -Path $ModulePath -Recurse -Force
        
        # Reinstall
        & "$PSScriptRoot\Install-OutlookSearch.ps1" -Force
        
        Write-Host "`n✓ Update complete! Backup saved to: $backupPath" -ForegroundColor Green
    }
    
} catch {
    Write-Error "Failed to check for updates: $_"
    Write-Host "You can manually reinstall using Install-OutlookSearch.ps1 -Force"
}
