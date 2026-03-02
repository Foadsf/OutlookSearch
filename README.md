# OutlookSearch

A production-grade PowerShell 7.4+ module that provides both a powerful CLI and an elegant TUI for advanced Outlook email searching with boolean logic, date ranges, and Markdown/PDF export capabilities using only built-in Windows technologies and the official ConsoleGuiTools module.

## Installation & Quick Start

### Prerequisites

- Windows 10/11 with Microsoft Outlook (2016-2026) installed
- PowerShell 7.4 or higher
- Terminal.Gui support via `Microsoft.PowerShell.ConsoleGuiTools`

### One-Line Installation

```powershell
# Run this in PowerShell 7.4+ as Administrator or current user
Invoke-RestMethod -Uri "https://raw.githubusercontent.com/Foadsf/OutlookSearch/master/Install-OutlookSearch.ps1" | Invoke-Expression
```

### Manual Installation

```powershell
# 1. Create module directory
$ModulePath = "$env:USERPROFILE\Documents\PowerShell\Modules\OutlookSearch"
New-Item -ItemType Directory -Path $ModulePath -Force

# 2. Save the module files
# 3. Import the module
Import-Module OutlookSearch -Force

# 4. Verify installation
Get-Command -Module OutlookSearch
```

### First Commands

```powershell
# Quick search
Search-Outlook -From "boss@company.com" -After "2025-01-01"

# Boolean logic
Search-Outlook -Subject "urgent" -Or @{Subject = "meeting"} -And @{From = "alice"} -Not @{HasAttachments = $true}

# Launch TUI
Search-Outlook -Interactive

# Export to Markdown
Search-Outlook -From "ceo@company.com" -ThisWeek -Markdown -ExportPath "C:\Reports\ceo-emails.md"
```

## License

Licensed under the GNU General Public License v3.0 (GPL-3.0). See the [LICENSE](LICENSE) file for details.
