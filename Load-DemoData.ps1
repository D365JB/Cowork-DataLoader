<#
.SYNOPSIS
    Copilot Cowork Demo Data Loader
.DESCRIPTION
    Populates a Microsoft 365 tenant with demo data for Copilot Cowork demos.
    Supports loading emails, calendar events, and OneDrive files independently
    or all at once.

    Data is defined in JSON files under the data/ folder and can be customized
    per demo scenario.

.PARAMETER DataTypes
    Which data to load: All, Emails, Calendar, Files (comma-separated).
    Default: All

.PARAMETER ConfigPath
    Path to config.json. Default: ./config.json

.PARAMETER WhatIf
    Preview what would be loaded without making changes.

.EXAMPLE
    .\Load-DemoData.ps1
    # Loads all demo data (emails, calendar, files)

.EXAMPLE
    .\Load-DemoData.ps1 -DataTypes Emails
    # Loads only emails

.EXAMPLE
    .\Load-DemoData.ps1 -DataTypes Calendar,Files
    # Loads calendar events and OneDrive files

.EXAMPLE
    .\Load-DemoData.ps1 -WhatIf
    # Preview mode - shows what would be loaded
#>

param(
    [ValidateSet("All", "Emails", "Calendar", "Files", "Chats", "SharePoint", "Skills", "Channels", "D365")]
    [string[]]$DataTypes = @("All"),

    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),

    [switch]$WhatIf
)

# ── Bootstrap ────────────────────────────────────────────────────────────────

$ErrorActionPreference = "Stop"
$scriptRoot = $PSScriptRoot
$dataDir    = Join-Path $scriptRoot "data"

# Load modules
$modulesDir = Join-Path $scriptRoot "modules"
. (Join-Path $modulesDir "Connect-DemoGraph.ps1")
. (Join-Path $modulesDir "Send-DemoEmails.ps1")
. (Join-Path $modulesDir "New-DemoCalendarEvents.ps1")
. (Join-Path $modulesDir "Upload-DemoFiles.ps1")
. (Join-Path $modulesDir "Send-DemoChats.ps1")
. (Join-Path $modulesDir "Initialize-DemoSharePoint.ps1")
. (Join-Path $modulesDir "Deploy-CoworkSkills.ps1")
. (Join-Path $modulesDir "Send-DemoChannelMessages.ps1")
. (Join-Path $modulesDir "Connect-DemoDataverse.ps1")
. (Join-Path $modulesDir "Initialize-DemoD365.ps1")

# Load config
if (-not (Test-Path $ConfigPath)) {
    Write-Host "[ERROR] Config file not found: $ConfigPath" -ForegroundColor Red
    exit 1
}

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

$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json | ConvertTo-Hashtable

# Resolve "All"
if ($DataTypes -contains "All") {
    $DataTypes = @("Emails", "Calendar", "Files", "Chats", "SharePoint", "Skills", "Channels", "D365")
}

# ── Banner ───────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║       Copilot Cowork - Demo Data Loader             ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Tenant:    $($config.tenant.domain)" -ForegroundColor White
Write-Host "  Company:   $($config.demo.company)" -ForegroundColor White
Write-Host "  Demo week: $($config.demo.weekStart)" -ForegroundColor White
Write-Host "  Loading:   $($DataTypes -join ', ')" -ForegroundColor White
Write-Host ""

# ── Load data files ──────────────────────────────────────────────────────────

$emails = @()
$events = @()
$files  = @()
$chats  = @()
$spFiles = @()

if ($DataTypes -contains "Emails") {
    $emailsPath = Join-Path $dataDir "emails.json"
    if (Test-Path $emailsPath) {
        $emails = Get-Content $emailsPath -Raw | ConvertFrom-Json
        Write-Host "  Loaded $($emails.Count) emails from emails.json" -ForegroundColor DarkGray
    } else {
        Write-Host "  [WARN] emails.json not found - skipping" -ForegroundColor Yellow
    }
}

if ($DataTypes -contains "Calendar") {
    $eventsPath = Join-Path $dataDir "calendar-events.json"
    if (Test-Path $eventsPath) {
        $events = Get-Content $eventsPath -Raw | ConvertFrom-Json
        Write-Host "  Loaded $($events.Count) calendar events from calendar-events.json" -ForegroundColor DarkGray
    } else {
        Write-Host "  [WARN] calendar-events.json not found - skipping" -ForegroundColor Yellow
    }
}

if ($DataTypes -contains "Files") {
    $filesPath = Join-Path $dataDir "files.json"
    if (Test-Path $filesPath) {
        $files = Get-Content $filesPath -Raw | ConvertFrom-Json
        Write-Host "  Loaded $($files.Count) files from files.json" -ForegroundColor DarkGray
    } else {
        Write-Host "  [WARN] files.json not found - skipping" -ForegroundColor Yellow
    }
}

if ($DataTypes -contains "Chats") {
    $chatsPath = Join-Path $dataDir "chats.json"
    if (Test-Path $chatsPath) {
        $chats = Get-Content $chatsPath -Raw | ConvertFrom-Json
        Write-Host "  Loaded $($chats.Count) chat messages from chats.json" -ForegroundColor DarkGray
    } else {
        Write-Host "  [WARN] chats.json not found - skipping" -ForegroundColor Yellow
    }
}

if ($DataTypes -contains "SharePoint") {
    $spPath = Join-Path $dataDir "sharepoint-files.json"
    if (Test-Path $spPath) {
        $spFiles = Get-Content $spPath -Raw | ConvertFrom-Json
        Write-Host "  Loaded $($spFiles.Count) SharePoint files from sharepoint-files.json" -ForegroundColor DarkGray
    } else {
        Write-Host "  [WARN] sharepoint-files.json not found - skipping" -ForegroundColor Yellow
    }
}

$channelMessages = @()
if ($DataTypes -contains "Channels") {
    $chPath = Join-Path $dataDir "channel-messages.json"
    if (Test-Path $chPath) {
        $channelMessages = Get-Content $chPath -Raw | ConvertFrom-Json
        Write-Host "  Loaded $($channelMessages.Count) channel messages from channel-messages.json" -ForegroundColor DarkGray
    } else {
        Write-Host "  [WARN] channel-messages.json not found - skipping" -ForegroundColor Yellow
    }
}

$skills = @()
if ($DataTypes -contains "Skills") {
    $skillsPath = Join-Path $dataDir "skills.json"
    if (Test-Path $skillsPath) {
        $skills = Get-Content $skillsPath -Raw | ConvertFrom-Json
        Write-Host "  Loaded $($skills.Count) Cowork skills from skills.json" -ForegroundColor DarkGray
    } else {
        Write-Host "  [WARN] skills.json not found - skipping" -ForegroundColor Yellow
    }
}

$d365Records = $null
if ($DataTypes -contains "D365") {
    $d365Path = Join-Path $dataDir "d365-records.json"
    if (Test-Path $d365Path) {
        $d365Records = Get-Content $d365Path -Raw | ConvertFrom-Json
        $acctCount = @($d365Records.accounts).Count
        $contactCount = @($d365Records.contacts).Count
        $oppCount = @($d365Records.opportunities).Count
        Write-Host "  Loaded D365 records: $acctCount accounts, $contactCount contacts, $oppCount opportunities" -ForegroundColor DarkGray
    } else {
        Write-Host "  [WARN] d365-records.json not found - skipping" -ForegroundColor Yellow
    }
}

Write-Host ""

# ── WhatIf preview ───────────────────────────────────────────────────────────

if ($WhatIf) {
    Write-Host "═══ WHATIF MODE - No changes will be made ═══" -ForegroundColor Yellow
    Write-Host ""

    if ($emails.Count -gt 0) {
        Write-Host "EMAILS ($($emails.Count)):" -ForegroundColor Cyan
        foreach ($e in $emails) {
            $fromName = $config.users[$e.from].displayName
            $toName   = $config.users[$e.to].displayName
            Write-Host "  [$($e.scenario)] $fromName -> $toName : $($e.subject)" -ForegroundColor White
        }
        Write-Host ""
    }

    if ($events.Count -gt 0) {
        $weekStart = [datetime]::Parse($config.demo.weekStart)
        Write-Host "CALENDAR EVENTS ($($events.Count)):" -ForegroundColor Cyan
        foreach ($ev in $events) {
            $d = $weekStart.AddDays($ev.dayOffset).ToString("ddd MM/dd")
            $t = if ($ev.allDay) { "all day" } else { "$($ev.startTime)-$($ev.endTime)" }
            Write-Host "  $d $t - $($ev.subject)" -ForegroundColor White
        }
        Write-Host ""
    }

    if ($files.Count -gt 0) {
        Write-Host "ONEDRIVE FILES ($($files.Count)):" -ForegroundColor Cyan
        foreach ($f in $files) {
            Write-Host "  $($f.remotePath) -> $($config.users[$f.owner].displayName)" -ForegroundColor White
        }
        Write-Host ""
    }

    if ($chats.Count -gt 0) {
        Write-Host "TEAMS CHATS ($($chats.Count) messages):" -ForegroundColor Cyan
        foreach ($c in $chats) {
            $fromName = $config.users[$c.from].displayName
            $toName   = $config.users[$c.to].displayName
            Write-Host "  $fromName -> $toName : $($c.topic)" -ForegroundColor White
        }
        Write-Host ""
    }

    if ($spFiles.Count -gt 0) {
        Write-Host "SHAREPOINT FILES ($($spFiles.Count)):" -ForegroundColor Cyan
        foreach ($sp in $spFiles) {
            Write-Host "  $($sp.remotePath) -> ApexSalesTeam site" -ForegroundColor White
        }
        Write-Host ""
    }

    if ($channelMessages.Count -gt 0) {
        $channels = $channelMessages | Select-Object -ExpandProperty channel -Unique
        Write-Host "TEAMS CHANNELS ($($channelMessages.Count) messages across $($channels.Count) channels):" -ForegroundColor Cyan
        foreach ($ch in $channels) {
            $chMsgs = @($channelMessages | Where-Object { $_.channel -eq $ch })
            Write-Host "  #$ch ($($chMsgs.Count) messages)" -ForegroundColor White
        }
        Write-Host ""
    }

    if ($skills.Count -gt 0) {
        Write-Host "COWORK SKILLS ($($skills.Count)):" -ForegroundColor Cyan
        foreach ($sk in $skills) {
            Write-Host "  $($sk.skillName) -> $($config.users[$sk.owner].displayName)'s OneDrive" -ForegroundColor White
        }
        Write-Host ""
    }

    if ($d365Records) {
        $acctCount = @($d365Records.accounts).Count
        $contactCount = @($d365Records.contacts).Count
        $oppCount = @($d365Records.opportunities).Count
        Write-Host "D365 RECORDS ($acctCount accounts, $contactCount contacts, $oppCount opportunities):" -ForegroundColor Cyan
        foreach ($acct in $d365Records.accounts) {
            Write-Host "  Account: $($acct.name) - `$$([math]::Round($acct.revenue / 1000000))M revenue" -ForegroundColor White
        }
        foreach ($contact in $d365Records.contacts) {
            Write-Host "  Contact: $($contact.firstName) $($contact.lastName) ($($contact.jobTitle))" -ForegroundColor White
        }
        foreach ($opp in $d365Records.opportunities) {
            Write-Host "  Opportunity: $($opp.name) - `$$([math]::Round($opp.amount / 1000000, 1))M ($($opp.stage))" -ForegroundColor White
        }
        Write-Host ""
    }

    Write-Host "Run without -WhatIf to execute." -ForegroundColor Yellow
    exit 0
}

# ── Ensure Graph modules ────────────────────────────────────────────────────

$requiredModules = @("Microsoft.Graph.Authentication")
foreach ($mod in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "Installing $mod..." -ForegroundColor Yellow
        Install-Module $mod -Scope CurrentUser -Force -AllowClobber
    }
}

# ── Execute: Emails (AppOnly auth) ───────────────────────────────────────────

if ($emails.Count -gt 0) {
    Write-Host "────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "EMAILS (app-only auth)" -ForegroundColor Cyan
    Write-Host "────────────────────────────────────────────" -ForegroundColor DarkGray

    $connected = Connect-DemoGraphAppOnly -Config $config
    if ($connected) {
        Write-Host ""
        Write-Host "EMAILS:" -ForegroundColor Cyan
        Send-DemoEmails -Config $config -Emails $emails
    } else {
        Write-Host "[SKIP] Emails skipped - could not authenticate." -ForegroundColor Yellow
    }
    Write-Host ""
}

# ── Execute: Chats (AppOnly auth - Migration API) ────────────────────────────

if ($chats.Count -gt 0) {
    Write-Host "────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "TEAMS CHATS (app-only auth - migration API)" -ForegroundColor Cyan
    Write-Host "────────────────────────────────────────────" -ForegroundColor DarkGray

    $connected = Connect-DemoGraphAppOnly -Config $config
    if ($connected) {
        Write-Host ""
        Write-Host "TEAMS CHATS:" -ForegroundColor Cyan
        Send-DemoChats -Config $config -Chats $chats
    } else {
        Write-Host "[SKIP] Chats skipped - could not authenticate." -ForegroundColor Yellow
    }
    Write-Host ""
}

# ── Execute: Calendar + Files + SharePoint (Delegated auth) ─────────────────

if ($events.Count -gt 0 -or $files.Count -gt 0 -or $spFiles.Count -gt 0 -or $skills.Count -gt 0 -or $channelMessages.Count -gt 0) {
    Write-Host "────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "CALENDAR, FILES & SHAREPOINT (delegated auth)" -ForegroundColor Cyan
    Write-Host "────────────────────────────────────────────" -ForegroundColor DarkGray

    $scopes = @()
    if ($events.Count -gt 0)   { $scopes += "Calendars.ReadWrite" }
    if ($files.Count -gt 0)    { $scopes += "Files.ReadWrite.All" }
    if ($skills.Count -gt 0)   { $scopes += "Files.ReadWrite.All" }
    if ($channelMessages.Count -gt 0) { $scopes += "Group.ReadWrite.All"; $scopes += "Channel.Create"; $scopes += "ChannelMessage.Send"; $scopes += "ChannelMessage.Read.All" }
    if ($spFiles.Count -gt 0)  { $scopes += "Sites.ReadWrite.All"; $scopes += "Group.ReadWrite.All" }
    $scopes += "User.Read.All"

    $connected = Connect-DemoGraphDelegated -Config $config -Scopes $scopes
    if ($connected) {
        if ($events.Count -gt 0) {
            Write-Host ""
            Write-Host "CALENDAR EVENTS:" -ForegroundColor Cyan
            New-DemoCalendarEvents -Config $config -Events $events
        }
        if ($files.Count -gt 0) {
            Write-Host ""
            Write-Host "ONEDRIVE FILES:" -ForegroundColor Cyan
            Upload-DemoFiles -Config $config -Files $files -DataDir $dataDir
        }
        if ($spFiles.Count -gt 0) {
            Write-Host ""
            Write-Host "SHAREPOINT SITE & FILES:" -ForegroundColor Cyan
            Initialize-DemoSharePoint -Config $config -SharePointFiles $spFiles -DataDir $dataDir
        }
        if ($skills.Count -gt 0) {
            Write-Host ""
            Write-Host "COWORK SKILLS:" -ForegroundColor Cyan
            Deploy-CoworkSkills -Config $config -Skills $skills -DataDir $dataDir
        }
        if ($channelMessages.Count -gt 0) {
            Write-Host ""
            Write-Host "TEAMS CHANNELS:" -ForegroundColor Cyan
            Send-DemoChannelMessages -Config $config -ChannelMessages $channelMessages
        }
    } else {
        Write-Host "[SKIP] Calendar/Files/SharePoint/Skills/Channels skipped - could not authenticate." -ForegroundColor Yellow
    }
}

# ── Execute: D365 / Dataverse (AppOnly auth) ────────────────────────────────

if ($d365Records) {
    Write-Host ""
    Write-Host "═══ D365 SALES RECORDS ═══" -ForegroundColor Cyan

    $dvConnection = Connect-DemoDataverse -Config $config
    if ($dvConnection) {
        Initialize-DemoD365 -Config $config -Connection $dvConnection -D365Records $d365Records
    } else {
        Write-Host "[SKIP] D365 records skipped - could not connect to Dataverse." -ForegroundColor Yellow
        Write-Host "  Ensure config.json has dataverse.environmentUrl set and the app registration" -ForegroundColor Yellow
        Write-Host "  has Dynamics CRM user_impersonation permission." -ForegroundColor Yellow
    }
}

# ── Done ─────────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "╔══════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║       Demo data loading complete!                   ║" -ForegroundColor Green
Write-Host "╚══════════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""
