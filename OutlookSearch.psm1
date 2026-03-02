#requires -Version 7.4
#requires -PSEdition Core

#region Initialization
$script:OutlookApp = $null
$script:Namespace = $null
$script:DefaultFolder = $null

function Initialize-OutlookConnection {
    [CmdletBinding()]
    param()
    
    try {
        if (-not $script:OutlookApp) {
            Write-Verbose "Initializing Outlook COM connection..."
            $script:OutlookApp = New-Object -ComObject Outlook.Application
            $script:Namespace = $script:OutlookApp.GetNamespace("MAPI")
            
            # Ensure Outlook is running (starts in background if needed)
            try {
                $null = $script:OutlookApp.ActiveExplorer()
            }
            catch {
                # Outlook wasn't running, start it minimized
                $proc = Start-Process -FilePath "outlook.exe" -WindowStyle Minimized -PassThru
                Start-Sleep -Seconds 2
                # Reconnect
                $script:OutlookApp = New-Object -ComObject Outlook.Application
                $script:Namespace = $script:OutlookApp.GetNamespace("MAPI")
            }
        }
        return $true
    }
    catch {
        Write-Error "Failed to connect to Outlook: $_"
        return $false
    }
}

function Get-OutlookFolder {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$FolderPath = "Inbox",
        
        [Parameter()]
        [int]$DefaultFolderType = 6  # 6 = Inbox
    )
    
    Initialize-OutlookConnection | Out-Null
    
    if ($FolderPath -eq "Inbox") {
        return $script:Namespace.GetDefaultFolder($DefaultFolderType)
    }
    
    # Handle custom folder paths like "Inbox\Projects\Important"
    $folders = $FolderPath -split '\\'
    $currentFolder = $script:Namespace.GetDefaultFolder(6)  # Start at Inbox
    
    foreach ($folderName in $folders) {
        $found = $false
        foreach ($subFolder in $currentFolder.Folders) {
            if ($subFolder.Name -eq $folderName) {
                $currentFolder = $subFolder
                $found = $true
                break
            }
        }
        if (-not $found) {
            throw "Folder '$folderName' not found in path '$FolderPath'"
        }
    }
    
    return $currentFolder
}

# Define default display properties for our custom email objects
$typeData = @{
    TypeName   = 'OutlookSearch.Email'
    MemberType = 'ScriptProperty'
    MemberName = 'SizeKB'
    Value      = { if ($this.Size) { [math]::Round($this.Size / 1KB, 2) } else { 0 } }
    Force      = $true
}
Update-TypeData @typeData -ErrorAction SilentlyContinue

$typeSet = @{
    TypeName                  = 'OutlookSearch.Email'
    DefaultDisplayPropertySet = 'ReceivedTime', 'SenderName', 'Subject', 'SizeKB'
    Force                     = $true
}
Update-TypeData @typeSet -ErrorAction SilentlyContinue

#endregion

#region Filter Builder
function Build-DaslFilter {
    <#
    .SYNOPSIS
    Builds an Outlook DASL filter string from a criteria hashtable.
    Uses DASL syntax exclusively — JET LIKE/property syntax is unreliable
    across Outlook versions and locale settings.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [hashtable]$Criteria,
        
        [Parameter()]
        [string]$RawFilter
    )
    
    if ($RawFilter) {
        return $RawFilter
    }
    
    # DASL property URNs (all tested and verified against Outlook COM)
    $propSubject = '"urn:schemas:httpmail:subject"'
    $propBody = '"urn:schemas:httpmail:textdescription"'
    $propFrom = '"urn:schemas:httpmail:fromname"'
    $propFromAddr = '"urn:schemas:httpmail:fromemail"'
    $propTo = '"urn:schemas:httpmail:displayto"'
    $propDate = '"urn:schemas:httpmail:datereceived"'
    $propRead = '"urn:schemas:httpmail:read"'
    $propImport = '"urn:schemas:httpmail:importance"'
    $propHasAttach = '"urn:schemas:httpmail:hasattachment"'
    
    $conditions = @()
    
    # Sender/From (partial match)
    if ($Criteria.From) {
        $escaped = $Criteria.From -replace "'", "''"
        $conditions += "($propFrom LIKE '%$escaped%' OR $propFromAddr LIKE '%$escaped%')"
    }
    
    # Recipients/To (partial match)
    if ($Criteria.To) {
        $escaped = $Criteria.To -replace "'", "''"
        $conditions += "$propTo LIKE '%$escaped%'"
    }
    
    # Subject
    if ($Criteria.Subject) {
        $escaped = $Criteria.Subject -replace "'", "''"
        if ($Criteria.SubjectExact) {
            $conditions += "$propSubject = '$escaped'"
        }
        else {
            $conditions += "$propSubject LIKE '%$escaped%'"
        }
    }
    
    # Body (content search)
    if ($Criteria.Body) {
        $escaped = $Criteria.Body -replace "'", "''"
        $conditions += "$propBody LIKE '%$escaped%'"
    }
    
    # Date ranges (ISO 8601 format works reliably with DASL)
    if ($Criteria.After) {
        $dateStr = (Get-Date $Criteria.After).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $conditions += "$propDate >= '$dateStr'"
    }
    if ($Criteria.Before) {
        $dateStr = (Get-Date $Criteria.Before).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $conditions += "$propDate <= '$dateStr'"
    }
    
    # Boolean properties
    if ($Criteria.HasAttachments) {
        $conditions += "$propHasAttach = 1"
    }
    if ($null -ne $Criteria.IsRead) {
        if ($Criteria.IsRead) {
            $conditions += "$propRead = 1"
        }
        else {
            $conditions += "$propRead = 0"
        }
    }
    if ($Criteria.IsFlagged) {
        # FlagStatus: 0=None, 1=Complete, 2=Flagged
        $conditions += '"http://schemas.microsoft.com/mapi/proptag/0x10900003" = 2'
    }
    
    # Importance (0=Low, 1=Normal, 2=High)
    if ($Criteria.Importance) {
        $impMap = @{ Low = 0; Normal = 1; High = 2 }
        $conditions += "$propImport = $($impMap[$Criteria.Importance])"
    }
    
    # Combine with AND logic (default)
    if ($conditions.Count -eq 0) {
        return ""
    }
    
    # Wrap everything in @SQL= prefix for DASL
    $combined = if ($conditions.Count -eq 1) {
        $conditions[0]
    }
    else {
        $conditions -join " AND "
    }
    
    return "@SQL=$combined"
}

function Build-ComplexFilter {
    <#
    .SYNOPSIS
    Combines multiple DASL filter parts with AND/OR/NOT boolean logic.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [array]$AndFilters,
        
        [Parameter()]
        [array]$OrFilters,
        
        [Parameter()]
        [array]$NotFilters
    )
    
    # Strip any existing @SQL= prefix from individual filters
    $stripPrefix = { param($f) if ($f -match '^@SQL=(.+)$') { $Matches[1] } else { $f } }
    
    $parts = @()
    
    foreach ($filter in $AndFilters) {
        if ($filter) { $parts += "($(& $stripPrefix $filter))" }
    }
    
    if ($OrFilters.Count -gt 0) {
        $orParts = @()
        foreach ($filter in $OrFilters) {
            if ($filter) { $orParts += $(& $stripPrefix $filter) }
        }
        if ($orParts.Count -gt 0) {
            $parts += "($($orParts -join ' OR '))"
        }
    }
    
    foreach ($filter in $NotFilters) {
        if ($filter) { $parts += "(NOT ($(& $stripPrefix $filter)))" }
    }
    
    if ($parts.Count -eq 0) { return "" }
    
    return "@SQL=$($parts -join ' AND ')"
}

#endregion

#region Core Search Function
function Search-Outlook {
    <#
    .SYNOPSIS
    Advanced Outlook email search with boolean logic and export capabilities.
    
    .DESCRIPTION
    Searches Outlook emails using JET/DASL query syntax with support for complex boolean logic,
    date ranges, sender/recipient filters, and multiple export formats.
    
    .PARAMETER From
    Sender name or email address (partial match supported)
    
    .PARAMETER To
    Recipient name or email address
    
    .PARAMETER Subject
    Subject line text to search
    
    .PARAMETER SubjectExact
    Exact subject match
    
    .PARAMETER Body
    Body content to search
    
    .PARAMETER After
    Received on or after this date
    
    .PARAMETER Before
    Received on or before this date
    
    .PARAMETER ThisWeek
    Received this calendar week
    
    .PARAMETER ThisMonth
    Received this calendar month
    
    .PARAMETER HasAttachments
    Only emails with attachments
    
    .PARAMETER IsRead
    Filter by read status
    
    .PARAMETER IsFlagged
    Only flagged emails
    
    .PARAMETER Importance
    Filter by importance: Low, Normal, High
    
    .PARAMETER Folder
    Outlook folder to search (default: Inbox)
    
    .PARAMETER And
    Hashtable of additional criteria to AND with main criteria
    
    .PARAMETER Or
    Hashtable of additional criteria to OR with main criteria
    
    .PARAMETER Not
    Hashtable of criteria to exclude (NOT)
    
    .PARAMETER RawFilter
    Raw JET/DASL filter string for advanced users
    
    .PARAMETER Interactive
    Launch TUI mode
    
    .PARAMETER Limit
    Maximum results to return
    
    .PARAMETER Markdown
    Export results as Markdown
    
    .PARAMETER PDF
    Export results as PDF
    
    .PARAMETER ExportPath
    Path for exported files
    
    .EXAMPLE
    Search-Outlook -From "boss@company.com" -After "2025-01-01"
    
    .EXAMPLE
    Search-Outlook -Subject "urgent" -Or @{Subject = "meeting"} -And @{From = "alice"}
    
    .EXAMPLE
    Search-Outlook -Interactive
    #>
    [CmdletBinding(DefaultParameterSetName = 'CLI')]
    [Alias('oso')]
    param(
        [Parameter(ParameterSetName = 'CLI')]
        [string]$From,
        
        [Parameter(ParameterSetName = 'CLI')]
        [string]$To,
        
        [Parameter(ParameterSetName = 'CLI')]
        [string]$Subject,
        
        [Parameter(ParameterSetName = 'CLI')]
        [switch]$SubjectExact,
        
        [Parameter(ParameterSetName = 'CLI')]
        [string]$Body,
        
        [Parameter(ParameterSetName = 'CLI')]
        [datetime]$After,
        
        [Parameter(ParameterSetName = 'CLI')]
        [datetime]$Before,
        
        [Parameter(ParameterSetName = 'CLI')]
        [switch]$ThisWeek,
        
        [Parameter(ParameterSetName = 'CLI')]
        [switch]$ThisMonth,
        
        [Parameter(ParameterSetName = 'CLI')]
        [switch]$HasAttachments,
        
        [Parameter(ParameterSetName = 'CLI')]
        [System.Nullable[bool]]$IsRead,
        
        [Parameter(ParameterSetName = 'CLI')]
        [switch]$IsFlagged,
        
        [Parameter(ParameterSetName = 'CLI')]
        [ValidateSet('Low', 'Normal', 'High')]
        [string]$Importance,
        
        [Parameter(ParameterSetName = 'CLI')]
        [string]$Folder = "Inbox",
        
        [Parameter(ParameterSetName = 'CLI')]
        [hashtable]$And,
        
        [Parameter(ParameterSetName = 'CLI')]
        [hashtable]$Or,
        
        [Parameter(ParameterSetName = 'CLI')]
        [hashtable]$Not,
        
        [Parameter(ParameterSetName = 'CLI')]
        [string]$RawFilter,
        
        [Parameter(ParameterSetName = 'TUI')]
        [switch]$Interactive,
        
        [Parameter(ParameterSetName = 'CLI')]
        [int]$Limit = 0,
        
        [Parameter(ParameterSetName = 'CLI')]
        [switch]$Markdown,
        
        [Parameter(ParameterSetName = 'CLI')]
        [switch]$PDF,
        
        [Parameter(ParameterSetName = 'CLI')]
        [string]$ExportPath,
        
        [Parameter(ValueFromRemainingArguments)]
        [string[]]$RemainingArgs
    )
    
    # Handle GNU-style or generic help arguments smoothly
    if ($RemainingArgs -match '^(?:--?(?:h|help|\?)|/h|/\?)$') {
        Get-Help Search-Outlook -Full
        return
    }
    
    # TUI Mode
    if ($Interactive) {
        Show-OutlookSearchTUI
        return
    }
    
    # CLI Mode
    try {
        Initialize-OutlookConnection | Out-Null
        
        $targetFolder = Get-OutlookFolder -FolderPath $Folder
        
        # Build criteria hashtable
        $criteria = @{
            From           = $From
            To             = $To
            Subject        = $Subject
            SubjectExact   = $SubjectExact
            Body           = $Body
            HasAttachments = $HasAttachments
            IsRead         = $IsRead
            IsFlagged      = $IsFlagged
            Importance     = $Importance
        }
        
        # Handle relative dates
        if ($ThisWeek) {
            $criteria.After = (Get-Date).AddDays( - ([int](Get-Date).DayOfWeek))
        }
        if ($ThisMonth) {
            $criteria.After = Get-Date -Day 1 -Hour 0 -Minute 0 -Second 0
        }
        if ($After) { $criteria.After = $After }
        if ($Before) { $criteria.Before = $Before }
        
        # Build main filter
        $mainFilter = Build-DaslFilter -Criteria $criteria -RawFilter $RawFilter
        
        # Handle complex boolean logic
        $andFilters = @()
        $orFilters = @()
        $notFilters = @()
        
        if ($mainFilter) { $andFilters += $mainFilter }
        
        if ($And) {
            $andFilters += Build-DaslFilter -Criteria $And
        }
        if ($Or) {
            $orFilters += Build-DaslFilter -Criteria $Or
        }
        if ($Not) {
            $notFilters += Build-DaslFilter -Criteria $Not
        }
        
        $finalFilter = Build-ComplexFilter -AndFilters $andFilters -OrFilters $orFilters -NotFilters $notFilters
        
        Write-Verbose "Final DASL Filter: $finalFilter"
        
        # Execute search
        $items = $targetFolder.Items
        $items.Sort("[ReceivedTime]", $true)  # Sort supports only JET properties
        
        if ($finalFilter) {
            $results = $items.Restrict($finalFilter)
        }
        else {
            $results = $items
        }
        
        # Convert to PowerShell objects
        $emails = @()
        $count = 0
        
        foreach ($item in $results) {
            if ($item -is [System.__ComObject]) {
                try {
                    $email = [PSCustomObject]@{
                        EntryID         = $item.EntryID
                        Subject         = $item.Subject
                        SenderName      = $item.SenderName
                        SenderEmail     = $(try { $item.SenderEmailAddress } catch { "" })
                        To              = ($item.To -split ';' | ForEach-Object { $_.Trim() }) -join ", "
                        CC              = $(try { ($item.CC -split ';' | ForEach-Object { $_.Trim() }) -join ", " } catch { "" })
                        ReceivedTime    = $item.ReceivedTime
                        SentTime        = $(try { $item.SentOn } catch { $null })
                        HasAttachments  = ($item.Attachments.Count -gt 0)
                        AttachmentCount = $item.Attachments.Count
                        IsRead          = -not $item.UnRead
                        IsFlagged       = ($item.FlagRequest -ne "")
                        Importance      = $(switch ($item.Importance) { 0 { "Low" } 2 { "High" } default { "Normal" } })
                        Size            = $(try { $item.Size } catch { 0 })
                        Body            = $item.Body
                        BodyHTML        = $(try { $item.HTMLBody } catch { "" })
                        Categories      = $(try { $item.Categories } catch { "" })
                        OutlookItem     = $item  # Keep reference for export
                    }
                    $email.PSTypeNames.Insert(0, 'OutlookSearch.Email')
                    
                    $emails += $email
                    $count++
                    
                    if ($Limit -gt 0 -and $count -ge $Limit) {
                        break
                    }
                }
                catch {
                    Write-Warning "Error processing email: $_"
                }
            }
        }
        
        # Output or Export
        if ($Markdown -or $PDF) {
            if (-not $ExportPath) {
                $ExportPath = "OutlookExport_$(Get-Date -Format 'yyyyMMdd_HHmmss').md"
            }
            
            Export-OutlookEmail -Emails $emails -Path $ExportPath -Format $(if ($PDF) { "PDF" } else { "Markdown" })
        }
        else {
            # Print a visually appealing summary to the host (does not pollute the pipeline)
            if ($emails.Count -gt 0) {
                Write-Host "`n✓ Found $($emails.Count) matching email(s)" -ForegroundColor Green -BackgroundColor Black
                Write-Host "=====================================================" -ForegroundColor DarkGray
            }
            else {
                Write-Host "`n✗ No matching emails found." -ForegroundColor Yellow
            }
        }
        
        # Return pure objects to the pipeline (automatically formatted as a table by TypeData)
        return $emails
        
    }
    catch {
        Write-Error "Search failed: $_"
    }
}

#endregion

#region Export Functions
function Export-OutlookEmail {
    <#
    .SYNOPSIS
    Exports Outlook emails to Markdown or PDF format.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [array]$Emails,
        
        [Parameter(Mandatory)]
        [string]$Path,
        
        [Parameter()]
        [ValidateSet('Markdown', 'PDF', 'HTML', 'EML')]
        [string]$Format = 'Markdown'
    )
    
    begin {
        $allEmails = @()
    }
    
    process {
        $allEmails += $Emails
    }
    
    end {
        if ($Format -eq 'Markdown') {
            $markdown = @"
# Outlook Email Export
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Total Emails: $($allEmails.Count)

---

"@
            
            foreach ($email in $allEmails) {
                $markdown += @"

## $($email.Subject)

**From:** $($email.SenderName) <$($email.SenderEmail)>  
**To:** $($email.To)  
**CC:** $($email.CC)  
**Date:** $($email.ReceivedTime)  
**Importance:** $($email.Importance)  
**Attachments:** $($email.AttachmentCount)

### Body

$($email.Body)

---

"@
            }
            
            $markdown | Out-File -FilePath $Path -Encoding UTF8
            Write-Host "Exported to Markdown: $Path" -ForegroundColor Green
        }
        elseif ($Format -eq 'PDF') {
            # Check for Pandoc
            $pandoc = Get-Command pandoc -ErrorAction SilentlyContinue
            
            if ($pandoc) {
                $mdPath = [System.IO.Path]::ChangeExtension($Path, ".md")
                Export-OutlookEmail -Emails $allEmails -Path $mdPath -Format Markdown
                
                & pandoc $mdPath -o $Path --pdf-engine=xelatex -V geometry:margin=1in
                Write-Host "Exported to PDF: $Path" -ForegroundColor Green
                
                Remove-Item $mdPath -ErrorAction SilentlyContinue
            }
            else {
                Write-Warning "Pandoc not found. Install with: winget install --source winget --exact --id JohnMacFarlane.Pandoc"
                Write-Host "Falling back to Markdown export..." -ForegroundColor Yellow
                $mdPath = [System.IO.Path]::ChangeExtension($Path, ".md")
                Export-OutlookEmail -Emails $allEmails -Path $mdPath -Format Markdown
                Write-Host "To convert to PDF manually, use Microsoft Print to PDF or install Pandoc" -ForegroundColor Cyan
            }
        }
        elseif ($Format -eq 'HTML') {
            $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Outlook Export</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
        .email { border: 1px solid #ddd; margin-bottom: 20px; padding: 15px; border-radius: 5px; }
        .header { background: #f5f5f5; padding: 10px; margin-bottom: 10px; }
        .subject { font-size: 1.2em; font-weight: bold; }
        .meta { color: #666; font-size: 0.9em; }
        .body { margin-top: 15px; white-space: pre-wrap; }
    </style>
</head>
<body>
    <h1>Outlook Email Export</h1>
    <p>Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
    <p>Total: $($allEmails.Count) emails</p>
    <hr>
"@
            
            foreach ($email in $allEmails) {
                $bodyHtml = [System.Web.HttpUtility]::HtmlEncode($email.Body) -replace "`n", "<br>"
                $html += @"
    <div class="email">
        <div class="header">
            <div class="subject">$([System.Web.HttpUtility]::HtmlEncode($email.Subject))</div>
            <div class="meta">
                From: $($email.SenderName) &lt;$($email.SenderEmail)&gt;<br>
                To: $($email.To)<br>
                Date: $($email.ReceivedTime)<br>
                Attachments: $($email.AttachmentCount)
            </div>
        </div>
        <div class="body">$bodyHtml</div>
    </div>
"@
            }
            
            $html += "</body></html>"
            $html | Out-File -FilePath $Path -Encoding UTF8
            Write-Host "Exported to HTML: $Path" -ForegroundColor Green
        }
    }
}

#endregion

#region TUI Implementation
function Show-OutlookSearchTUI {
    <#
    .SYNOPSIS
    Interactive Terminal UI for Outlook email search.
    #>
    [CmdletBinding()]
    param()
    
    # Check for ConsoleGuiTools
    $guiTools = Get-Module Microsoft.PowerShell.ConsoleGuiTools -ListAvailable
    
    if (-not $guiTools) {
        Write-Host "Installing required module: Microsoft.PowerShell.ConsoleGuiTools..." -ForegroundColor Yellow
        try {
            Install-Module Microsoft.PowerShell.ConsoleGuiTools -Force -Scope CurrentUser
            Import-Module Microsoft.PowerShell.ConsoleGuiTools
        }
        catch {
            Write-Error "Failed to install ConsoleGuiTools. Install manually: Install-Module Microsoft.PowerShell.ConsoleGuiTools -Force"
            return
        }
    }
    
    Import-Module Microsoft.PowerShell.ConsoleGuiTools
    
    # Initialize connection
    if (-not (Initialize-OutlookConnection)) {
        return
    }
    
    # Simple TUI using Out-ConsoleGridView for selection
    # For a full Terminal.Gui implementation, we'd need more complex code
    
    Write-Host @"
╔══════════════════════════════════════════════════════════════╗
║           OUTLOOK SEARCH - INTERACTIVE MODE                  ║
╠══════════════════════════════════════════════════════════════╣
║  Build your search query:                                    ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Cyan
    
    # Interactive input collection
    $searchParams = @{}
    
    Write-Host "Press Enter to skip any field`n" -ForegroundColor Gray
    
    $from = Read-Host "From (sender email/name)"
    if ($from) { $searchParams['From'] = $from }
    
    $subject = Read-Host "Subject (contains)"
    if ($subject) { $searchParams['Subject'] = $subject }
    
    Write-Host "`nDate Range:" -ForegroundColor Yellow
    Write-Host "1. Today"
    Write-Host "2. This Week"
    Write-Host "3. This Month"
    Write-Host "4. Custom Range"
    Write-Host "5. All Time"
    
    $dateChoice = Read-Host "`nSelect (1-5)"
    switch ($dateChoice) {
        '1' { $searchParams['After'] = (Get-Date).Date }
        '2' { $searchParams['ThisWeek'] = $true }
        '3' { $searchParams['ThisMonth'] = $true }
        '4' { 
            $after = Read-Host "After date (yyyy-MM-dd)"
            $searchParams['After'] = [datetime]$after
            $before = Read-Host "Before date (yyyy-MM-dd)"
            $searchParams['Before'] = [datetime]$before
        }
    }
    
    $hasAttach = Read-Host "`nHas attachments? (y/n/any)"
    if ($hasAttach -eq 'y') { $searchParams['HasAttachments'] = $true }
    if ($hasAttach -eq 'n') { $searchParams['HasAttachments'] = $false }
    
    Write-Host "`nSearching..." -ForegroundColor Green
    
    try {
        $results = Search-Outlook @searchParams
        
        if ($results.Count -eq 0) {
            Write-Host "`nNo emails found matching your criteria." -ForegroundColor Yellow
            return
        }
        
        # Use ConsoleGuiTools for selection
        $selected = $results | 
        Select-Object Subject, SenderName, ReceivedTime, @{N = "SizeKB"; E = { [math]::Round($_.Size / 1KB, 2) } }, EntryID |
        Out-ConsoleGridView -Title "Select emails (Space to multi-select, Enter to confirm)"
        
        if ($selected) {
            Write-Host "`nSelected $($selected.Count) email(s)" -ForegroundColor Green
            
            $action = Read-Host "`nAction: [V]iew, [E]xport Markdown, [P]DF, [S]ave .eml, [Q]uit"
            
            switch ($action.ToUpper()) {
                'V' {
                    $fullEmails = $results | Where-Object { $_.EntryID -in $selected.EntryID }
                    $fullEmails | Format-List Subject, SenderName, SenderEmail, To, ReceivedTime, Body
                }
                'E' {
                    $path = Read-Host "Export path (.md)"
                    $fullEmails = $results | Where-Object { $_.EntryID -in $selected.EntryID }
                    Export-OutlookEmail -Emails $fullEmails -Path $path -Format Markdown
                }
                'P' {
                    $path = Read-Host "Export path (.pdf)"
                    $fullEmails = $results | Where-Object { $_.EntryID -in $selected.EntryID }
                    Export-OutlookEmail -Emails $fullEmails -Path $path -Format PDF
                }
                'S' {
                    $folder = Read-Host "Save folder path"
                    if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
                    
                    $fullEmails = $results | Where-Object { $_.EntryID -in $selected.EntryID }
                    foreach ($email in $fullEmails) {
                        $safeSubject = $email.Subject -replace '[\\/:*?"<>|]', '_'
                        $fileName = "$folder\$($email.ReceivedTime.ToString('yyyyMMdd_HHmmss'))_$safeSubject.msg"
                        $email.OutlookItem.SaveAs($fileName, 3)  # 3 = olMSG
                        Write-Host "Saved: $fileName" -ForegroundColor Green
                    }
                }
            }
        }
        
    }
    catch {
        Write-Error "TUI operation failed: $_"
    }
}

#endregion

#region Aliases
New-Alias -Name oso -Value Search-Outlook -Force
#endregion
