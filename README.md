# Copilot Cowork - Demo Data Loader

Reusable PowerShell tool for populating a Microsoft 365 tenant with demo data for **Copilot Cowork** demos. Loads emails, calendar events, OneDrive files (including Office documents), Teams chats, Teams channel messages, SharePoint site content, custom Cowork skills, and Dynamics 365 records from JSON data files.

## Folder Structure

```
CoworkDataLoader/
├── Load-DemoData.ps1              # Main entry point
├── Reset-DemoData.ps1             # Cleanup script (remove demo data)
├── Setup-AppRegistration.ps1      # One-time Entra ID app setup
├── Generate-OfficeFiles.ps1       # One-time: generates .docx/.xlsx/.pptx files
├── config.json                    # Tenant, users, app credentials
├── DEMO-SCRIPT.doc                # Demo script with 12 prompts and talking points
├── modules/
│   ├── Connect-DemoGraph.ps1      # Auth helpers (AppOnly + Delegated)
│   ├── Send-DemoEmails.ps1        # Email sender (with backdated timestamps)
│   ├── New-DemoCalendarEvents.ps1 # Calendar event creator
│   ├── Upload-DemoFiles.ps1       # OneDrive file uploader (text + binary)
│   ├── Send-DemoChats.ps1         # Teams 1:1 chat sender
│   ├── Initialize-DemoSharePoint.ps1  # SharePoint site + doc library
│   ├── Deploy-CoworkSkills.ps1    # Custom Cowork skills uploader
│   └── Send-DemoChannelMessages.ps1   # Teams channel messages
└── data/
    ├── emails.json                # Email definitions (26 emails)
    ├── calendar-events.json       # Calendar events (24 events)
    ├── files.json                 # OneDrive file manifest (10 files)
    ├── chats.json                 # Teams chat messages (12 messages)
    ├── sharepoint-files.json      # SharePoint document manifest (10 files)
    ├── channel-messages.json      # Teams channel messages (15 messages)
    ├── skills.json                # Cowork skills manifest (2 skills)
    ├── skills/                    # Custom skill definitions
    │   ├── deal-review/SKILL.md   # Deal Review Brief skill
    │   └── weekly-status/SKILL.md # Weekly Executive Status Report skill
    └── files/                     # Local files to upload
        ├── Adatum Corp Briefing Deck.pptx
        ├── Adatum Corp Meeting Notes - Jan 15.docx
        ├── ProLine X Competitive Analysis.xlsx
        ├── ProLine X Launch Plan - Exec Summary.docx
        ├── ProLine X Pipeline Tracker.xlsx
        ├── Adatum Corp Meeting Notes - Jan 15.txt
        ├── Adatum Corp Briefing Deck - Draft.txt
        ├── ProLine X Positioning - Internal Draft.txt
        ├── ProLine X Battle Card - Internal.txt
        └── ProLine X Launch Plan - Exec Summary DRAFT.txt
```

## Prerequisites

- **PowerShell 5.1+** (Windows) or **PowerShell 7+** (cross-platform)
- **Microsoft.Graph.Authentication** module (auto-installed if missing)
- **Entra ID app registration** with application permissions (see Authentication table)
- **Global Admin** or equivalent to run the one-time setup

## Quick Start

### 1. First-time setup (once per tenant)

```powershell
.\Setup-AppRegistration.ps1
```

Creates an app registration with all required permissions and optionally updates `config.json`.

### 2. Edit config.json

Update the `users` section and set `weekStart` to the Monday of your demo week.

### 3. Load all demo data

```powershell
.\Load-DemoData.ps1
```

### 4. Load specific data types

```powershell
.\Load-DemoData.ps1 -DataTypes Emails
.\Load-DemoData.ps1 -DataTypes Calendar,Files
.\Load-DemoData.ps1 -DataTypes Chats,SharePoint
.\Load-DemoData.ps1 -DataTypes Skills,Channels
.\Load-DemoData.ps1 -WhatIf   # Preview mode
```

Valid data types: `All`, `Emails`, `Calendar`, `Files`, `Chats`, `SharePoint`, `Skills`, `Channels`, `D365`

### 5. Reset / Cleanup

```powershell
.\Reset-DemoData.ps1                       # Reset all data types
.\Reset-DemoData.ps1 -DataTypes Emails     # Reset only emails
.\Reset-DemoData.ps1 -WhatIf              # Preview what would be deleted
```

## Authentication

| Data Type | Auth Flow | Permissions | Why |
|-----------|-----------|-------------|-----|
| Emails | AppOnly (client creds) | `Mail.Send`, `Mail.ReadWrite` | Send as any user; backdate timestamps |
| Chats | AppOnly (client creds) | `Chat.Create`, `Chat.ReadWrite.All`, `Teamwork.Migrate.All` | Create chats, import messages as any user |
| Calendar | Delegated (interactive) | `Calendars.ReadWrite` | Events on signed-in user's calendar |
| Files | Delegated (interactive) | `Files.ReadWrite.All` | Upload to signed-in user's OneDrive |
| SharePoint | Delegated (interactive) | `Sites.ReadWrite.All`, `Group.ReadWrite.All` | Create team site + upload docs |
| Skills | Delegated (interactive) | `Files.ReadWrite.All` | Upload skill files to OneDrive |
| Channels | Delegated (interactive) | `Group.ReadWrite.All`, `Channel.Create`, `ChannelMessage.Send`, `ChannelMessage.Read.All` | Create team, channels, post messages; read for dedup |
| D365 | AppOnly (client creds) | Dataverse app user | Create accounts, contacts, opportunities |

## Email Backdating

Emails include `dayOffset` and `time` fields to create realistic arrival timestamps. The module places messages directly into the recipient's mailbox with the correct `receivedDateTime`, making the inbox look natural for demos. Requires `Mail.ReadWrite` (Application) permission — falls back to standard send if unavailable.

## Demo Prompts

| # | Prompt | Data Sources | Description |
|---|--------|-------------|-------------|
| 1 | **Catch Me Up** | Emails + Chats + Files + Calendar | Monday morning catch-up across all sources |
| 2 | **Deep Dive on the Deal** | Emails + CRM + Files + Chats | Full Adatum Corp deal picture ($2.4M) |
| 3 | **Meeting Prep** | Calendar + Emails + CRM | Prep brief for Thursday's Adatum meeting |
| 4 | **Research the Customer** | Emails + CRM | Adatum Corp strategic direction and intel |
| 5 | **Product Launch Status** | Emails + Calendar + Chats | ProLine X timeline, pipeline, competitive |
| 6 | **Draft a Deliverable** | Emails + Files + Chats | Board-ready executive summary from existing data |
| 7 | **Work with Office Documents** | OneDrive (.xlsx) | Read Pipeline Tracker spreadsheet |
| 8 | **Summarize a Presentation** | OneDrive (.pptx) | Summarize Adatum Briefing Deck |
| 9 | **Competitive Analysis** | OneDrive (.xlsx) + Emails | Build talking points from spreadsheet + emails |
| 10 | **Escalation Triage** | Emails + Channels | Identify urgent items and priorities |
| 11 | **Use a Custom Skill** | All sources (via skill) | Run Deal Review Brief skill for Adatum |
| 12 | **Channel Context** | Teams Channels | Summarize ProLine X Launch + Adatum Deal Room |

## Customizing Data

- **Add emails**: Edit `data/emails.json`. Set `dayOffset`/`time` for realistic timestamps. Set `isRead` per email.
- **Change demo week**: Update `config.json > demo > weekStart` (Monday of demo week).
- **Add files**: Place in `data/files/` and add entry to `data/files.json`. Supports .txt, .docx, .xlsx, .pptx.
- **Add Office files**: Run `Generate-OfficeFiles.ps1` to regenerate, or add your own to `data/files/`.
- **Add SharePoint docs**: Add entry to `data/sharepoint-files.json` (reuses files from `data/files/`).
- **Add chats**: Add entries to `data/chats.json` with `from`, `to`, `topic`, `message`.
- **Add channel messages**: Edit `data/channel-messages.json`. Set `channelName`, `from`, `dayOffset`.
- **Add skills**: Create `data/skills/{name}/SKILL.md` and add entry to `data/skills.json`.
- **New tenant**: Copy folder, update `config.json`, run `Setup-AppRegistration.ps1`.

## Idempotent Re-runs

The loader is safe to run multiple times. Each data type handles duplicates differently:

| Data Type | Strategy | Details |
|-----------|----------|---------|
| Emails | Delete + recreate | Searches recipient mailbox by subject; deletes old copy, creates new one |
| Calendar | Patch existing | Pre-fetches week's events; updates matching subjects via PATCH |
| Chats | Skip existing | Checks if conversation already has messages; skips if so |
| Channels | Skip existing | Reads channel messages; skips if message snippet already present |
| Files | Overwrite | Uploads replace existing files at the same path |
| SharePoint | Overwrite | Same as Files |
| Skills | Overwrite | Same as Files |

Re-running shows `[UPD]` for updated records and `[SKIP]` for skipped duplicates.

## Troubleshooting

| Issue | Fix |
|-------|-----|
| WAM login dialog hidden | Script sets `$env:MSAL_ENABLE_WAM = "0"` automatically |
| 403 on email send | Verify app has `Mail.Send` **application** permission with admin consent |
| 403 on chat send | Verify app has `Chat.Create` + `Chat.ReadWrite.All` with admin consent |
| Emails arrive with current timestamp | Add `Mail.ReadWrite` (Application) permission for backdating |
| SharePoint site creation fails | Ensure delegated consent for `Sites.ReadWrite.All` + `Group.ReadWrite.All` |
| `ConvertFrom-Json -AsHashtable` error | On PS 5.1, the script handles this automatically |
| Calendar events wrong dates | Check `weekStart` in config.json matches your demo week Monday |
| Teams chats can't be deleted | Graph API limitation — chats cannot be deleted programmatically |
| Teams channel creation fails | Ensure delegated consent for `Group.ReadWrite.All` + `Channel.Create` |
| Skills not visible in Cowork | Check OneDrive `/Documents/Cowork/skills/{name}/SKILL.md` path exists |
| Office files upload as 0 bytes | Binary files require `[System.IO.File]::ReadAllBytes()` — already handled |
| Team provisioning timeout | Teams can take 30-60s to provision after group creation; script retries automatically |
