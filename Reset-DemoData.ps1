<#
.SYNOPSIS
    Resets/cleans up all demo data from the tenant.
.DESCRIPTION
    Deletes emails, calendar events, OneDrive files, Teams chats, and SharePoint
    content created by Load-DemoData.ps1. Uses search + subject/name matching
    to find and remove only demo-generated items.

.PARAMETER DataTypes
    Which data to clean: All, Emails, Calendar, Files, Chats, SharePoint
    Default: All

.PARAMETER ConfigPath
    Path to config.json. Default: ./config.json

.PARAMETER WhatIf
    Preview what would be deleted without making changes.

.EXAMPLE
    .\Reset-DemoData.ps1
    # Deletes all demo data

.EXAMPLE
    .\Reset-DemoData.ps1 -DataTypes Emails,Calendar
    # Deletes only demo emails and calendar events

.EXAMPLE
    .\Reset-DemoData.ps1 -WhatIf
    # Preview mode
#>

param(
    [ValidateSet("All", "Emails", "Calendar", "Files", "Chats", "SharePoint")]
    [string[]]$DataTypes = @("All"),

    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),

    [switch]$WhatIf
)

$ErrorActionPreference = "Stop"
$scriptRoot = $PSScriptRoot
$dataDir    = Join-Path $scriptRoot "data"

# Helper: convert PSCustomObject to nested hashtable (PS 5.1 compat)
function ConvertTo-Hashtable {
    param([Parameter(ValueFromPipeline)]$InputObject)
    process {
        if ($InputObject -is [System.Collections.IDictionary]) { return $InputObject }
        if ($InputObject -is [PSCustomObject]) {
            $ht = @{}
            foreach ($prop in $InputObject.PSObject.Properties) {
                $ht[$prop.Name] = ConvertTo-Hashtable $prop.Value
            }
            return $ht
        }
        return $InputObject
    }
}

if (-not (Test-Path $ConfigPath)) {
    Write-Host "[ERROR] Config not found: $ConfigPath" -ForegroundColor Red
    exit 1
}
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json | ConvertTo-Hashtable

if ($DataTypes -contains "All") {
    $DataTypes = @("Emails", "Calendar", "Files", "Chats", "SharePoint")
}

Write-Host ""
Write-Host "╔══════════════════════════════════════════════════════╗" -ForegroundColor Red
Write-Host "║       Copilot Cowork - Demo Data RESET              ║" -ForegroundColor Red
Write-Host "╚══════════════════════════════════════════════════════╝" -ForegroundColor Red
Write-Host ""
Write-Host "  Tenant:   $($config.tenant.domain)" -ForegroundColor White
Write-Host "  Cleaning: $($DataTypes -join ', ')" -ForegroundColor White
Write-Host ""

if (-not $WhatIf) {
    $confirm = Read-Host "This will DELETE demo data. Type 'yes' to continue"
    if ($confirm -ne 'yes') {
        Write-Host "Aborted." -ForegroundColor Yellow
        exit 0
    }
}

# ── Load data files to know what to delete ───────────────────────────────────

# Build a set of known email subjects for matching
$emailSubjects = @()
$emailsPath = Join-Path $dataDir "emails.json"
if (Test-Path $emailsPath) {
    $emailData = Get-Content $emailsPath -Raw | ConvertFrom-Json
    $emailSubjects = $emailData | ForEach-Object { $_.subject }
}

# Build set of known calendar subjects
$calendarSubjects = @()
$calPath = Join-Path $dataDir "calendar-events.json"
if (Test-Path $calPath) {
    $calData = Get-Content $calPath -Raw | ConvertFrom-Json
    $calendarSubjects = $calData | ForEach-Object { $_.subject } | Sort-Object -Unique
}

# Build set of known file paths
$filePaths = @()
$filesPath = Join-Path $dataDir "files.json"
if (Test-Path $filesPath) {
    $fileData = Get-Content $filesPath -Raw | ConvertFrom-Json
    $filePaths = $fileData | ForEach-Object { $_.remotePath }
}

# ── Connect (AppOnly needed for email delete, delegated for the rest) ────────

. (Join-Path $scriptRoot "modules" "Connect-DemoGraph.ps1")

$jamesEmail = $config.users["james"].email

# ── Delete Emails ────────────────────────────────────────────────────────────

if ($DataTypes -contains "Emails" -and $emailSubjects.Count -gt 0) {
    Write-Host "────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "EMAILS - Deleting from inbox" -ForegroundColor Red
    Write-Host "────────────────────────────────────────────" -ForegroundColor DarkGray

    $connected = Connect-DemoGraphAppOnly -Config $config
    if ($connected) {
        $deleted = 0
        foreach ($subj in $emailSubjects) {
            try {
                $encodedSubject = [System.Uri]::EscapeDataString($subj)
                $filter = "subject eq '$($subj -replace "'", "''")'"
                $uri = "https://graph.microsoft.com/v1.0/users/$jamesEmail/messages?`$filter=$filter&`$select=id,subject&`$top=50"
                $result = Invoke-MgGraphRequest -Method GET -Uri $uri
                foreach ($msg in $result.value) {
                    if ($WhatIf) {
                        Write-Host "  [WOULD DELETE] $($msg.subject)" -ForegroundColor Yellow
                    } else {
                        Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/users/$jamesEmail/messages/$($msg.id)"
                        Write-Host "  [DELETED] $($msg.subject)" -ForegroundColor Green
                    }
                    $deleted++
                }
            } catch {
                Write-Host "  [FAIL] Could not delete '$subj': $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        Write-Host "[EMAILS] $deleted items $(if ($WhatIf) {'would be '})deleted." -ForegroundColor $(if ($deleted -gt 0) { 'Green' } else { 'DarkGray' })
    }
    Write-Host ""
}

# ── Delete Calendar Events ───────────────────────────────────────────────────

if ($DataTypes -contains "Calendar" -and $calendarSubjects.Count -gt 0) {
    Write-Host "────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "CALENDAR - Deleting events" -ForegroundColor Red
    Write-Host "────────────────────────────────────────────" -ForegroundColor DarkGray

    # Need delegated auth for calendar
    $connected = Connect-DemoGraphDelegated -Config $config -Scopes @("Calendars.ReadWrite")
    if ($connected) {
        $weekStart = [datetime]::Parse($config.demo.weekStart)
        $weekEnd   = $weekStart.AddDays(6)
        $startStr  = $weekStart.ToString("yyyy-MM-ddT00:00:00Z")
        $endStr    = $weekEnd.ToString("yyyy-MM-ddT23:59:59Z")

        $deleted = 0
        try {
            $uri = "https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=$startStr&endDateTime=$endStr&`$select=id,subject&`$top=100"
            $result = Invoke-MgGraphRequest -Method GET -Uri $uri
            foreach ($evt in $result.value) {
                if ($calendarSubjects -contains $evt.subject) {
                    if ($WhatIf) {
                        Write-Host "  [WOULD DELETE] $($evt.subject)" -ForegroundColor Yellow
                    } else {
                        Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/me/events/$($evt.id)"
                        Write-Host "  [DELETED] $($evt.subject)" -ForegroundColor Green
                    }
                    $deleted++
                }
            }
        } catch {
            Write-Host "  [FAIL] Calendar query error: $($_.Exception.Message)" -ForegroundColor Red
        }
        Write-Host "[CALENDAR] $deleted items $(if ($WhatIf) {'would be '})deleted." -ForegroundColor $(if ($deleted -gt 0) { 'Green' } else { 'DarkGray' })
    }
    Write-Host ""
}

# ── Delete OneDrive Files ────────────────────────────────────────────────────

if ($DataTypes -contains "Files" -and $filePaths.Count -gt 0) {
    Write-Host "────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "FILES - Deleting from OneDrive" -ForegroundColor Red
    Write-Host "────────────────────────────────────────────" -ForegroundColor DarkGray

    # Reuse delegated if still connected, otherwise reconnect
    $ctx = Get-MgContext
    if (-not $ctx -or $ctx.AuthType -eq 'AppOnly') {
        $connected = Connect-DemoGraphDelegated -Config $config -Scopes @("Files.ReadWrite.All")
    }

    $deleted = 0
    foreach ($remotePath in $filePaths) {
        try {
            $encodedPath = $remotePath -replace ' ', '%20'
            $uri = "https://graph.microsoft.com/v1.0/me/drive/root:/$encodedPath"
            if ($WhatIf) {
                # Check if it exists
                $item = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue
                Write-Host "  [WOULD DELETE] $remotePath" -ForegroundColor Yellow
                $deleted++
            } else {
                Invoke-MgGraphRequest -Method DELETE -Uri $uri
                Write-Host "  [DELETED] $remotePath" -ForegroundColor Green
                $deleted++
            }
        } catch {
            Write-Host "  [SKIP] $remotePath not found or already deleted" -ForegroundColor DarkGray
        }
    }
    Write-Host "[FILES] $deleted items $(if ($WhatIf) {'would be '})deleted." -ForegroundColor $(if ($deleted -gt 0) { 'Green' } else { 'DarkGray' })
    Write-Host ""
}

# ── Delete Teams Chats ───────────────────────────────────────────────────────

if ($DataTypes -contains "Chats") {
    Write-Host "────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "CHATS - Teams chat cleanup" -ForegroundColor DarkGray
    Write-Host "────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  [INFO] Teams chats cannot be deleted via Graph API." -ForegroundColor Yellow
    Write-Host "  Chat messages persist but will scroll out of view with normal usage." -ForegroundColor DarkGray
    Write-Host ""
}

# ── Delete SharePoint Content ────────────────────────────────────────────────

if ($DataTypes -contains "SharePoint") {
    Write-Host "────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "SHAREPOINT - Deleting site content" -ForegroundColor Red
    Write-Host "────────────────────────────────────────────" -ForegroundColor DarkGray

    $spSiteName = "ApexSalesTeam"
    $ctx = Get-MgContext
    if (-not $ctx -or $ctx.AuthType -eq 'AppOnly') {
        $connected = Connect-DemoGraphDelegated -Config $config -Scopes @("Sites.ReadWrite.All")
    }

    try {
        $domain = $config.tenant.domain -replace '\.OnMicrosoft\.com$', '.sharepoint.com'
        $siteUri = "https://graph.microsoft.com/v1.0/sites/${domain}:/sites/${spSiteName}"
        $site = Invoke-MgGraphRequest -Method GET -Uri $siteUri -ErrorAction Stop

        # Delete files from the default document library
        $driveUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/drive/root/children"
        $items = Invoke-MgGraphRequest -Method GET -Uri $driveUri

        $deleted = 0
        foreach ($item in $items.value) {
            if ($WhatIf) {
                Write-Host "  [WOULD DELETE] $($item.name)" -ForegroundColor Yellow
            } else {
                Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drive/items/$($item.id)"
                Write-Host "  [DELETED] $($item.name)" -ForegroundColor Green
            }
            $deleted++
        }
        Write-Host "[SHAREPOINT] $deleted items $(if ($WhatIf) {'would be '})deleted." -ForegroundColor $(if ($deleted -gt 0) { 'Green' } else { 'DarkGray' })
    } catch {
        Write-Host "  [SKIP] SharePoint site '$spSiteName' not found or not accessible" -ForegroundColor DarkGray
    }
    Write-Host ""
}

# ── Done ─────────────────────────────────────────────────────────────────────

Write-Host "╔══════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║       Reset complete!                               ║" -ForegroundColor Green
Write-Host "╚══════════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""
