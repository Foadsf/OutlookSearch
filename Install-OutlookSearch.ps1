<#
.SYNOPSIS
One-click installer for OutlookSearch PowerShell module.
#>
[CmdletBinding()]
param(
    [string]$InstallPath = "$env:USERPROFILE\Documents\PowerShell\Modules\OutlookSearch",
    [switch]$Force
)

Write-Host "OutlookSearch Module Installer" -ForegroundColor Cyan
Write-Host "==============================`n" -ForegroundColor Cyan

# Check PowerShell version
if ($PSVersionTable.PSVersion -lt [version]"7.4") {
    Write-Error "PowerShell 7.4 or higher required. Current: $($PSVersionTable.PSVersion)"
    Write-Host "Download from: https://github.com/PowerShell/PowerShell/releases"
    exit 1
}

# Check for Outlook
try {
    $null = New-Object -ComObject Outlook.Application
    Write-Host "✓ Microsoft Outlook detected" -ForegroundColor Green
}
catch {
    Write-Warning "Microsoft Outlook not detected. Module requires Outlook to be installed."
}

# Create directory
if (Test-Path $InstallPath) {
    if (-not $Force) {
        $overwrite = Read-Host "Module already exists. Overwrite? (Y/N)"
        if ($overwrite -ne 'Y') {
            Write-Host "Installation cancelled."
            exit 0
        }
    }
    Remove-Item -Path $InstallPath -Recurse -Force
}

New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null

# Download module files (in production, these would be from GitHub/raw URLs)
$baseUrl = "https://raw.githubusercontent.com/Foadsf/OutlookSearch/master"

$files = @(
    "OutlookSearch.psd1",
    "OutlookSearch.psm1"
)

foreach ($file in $files) {
    Write-Host "Downloading $file..." -NoNewline
    try {
        Invoke-RestMethod -Uri "$baseUrl/$file" -OutFile "$InstallPath\$file"
        Write-Host " ✓" -ForegroundColor Green
    }
    catch {
        Write-Host " ✗ Failed" -ForegroundColor Red
        Write-Error $_
        exit 1
    }
}

# Import module
Write-Host "`nImporting module..." -NoNewline
Import-Module $InstallPath -Force
Write-Host " ✓" -ForegroundColor Green

# Verify
$cmd = Get-Command Search-Outlook -ErrorAction SilentlyContinue
if ($cmd) {
    Write-Host "`n✓ Installation successful!" -ForegroundColor Green
    Write-Host "`nQuick start:"
    Write-Host "  Search-Outlook -From 'boss@company.com' -ThisWeek"
    Write-Host "  Search-Outlook -Interactive"
    Write-Host "  Get-Help Search-Outlook -Full"
}
else {
    Write-Error "Installation verification failed."
}
