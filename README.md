# Copilot Cowork - Demo Data Loader

Reusable PowerShell tool for populating a Microsoft 365 tenant with demo data for **Copilot Cowork** demos. Loads emails, calendar events, OneDrive files, Teams chats, and SharePoint site content from JSON data files.

## Folder Structure

```
CoworkDataLoader/
├── Load-DemoData.ps1              # Main entry point
├── Reset-DemoData.ps1             # Cleanup script (remove demo data)
├── Setup-AppRegistration.ps1      # One-time Entra ID app setup
├── config.json                    # Tenant, users, app credentials
├── modules/
│   ├── Connect-DemoGraph.ps1      # Auth helpers (AppOnly + Delegated)
│   ├── Send-DemoEmails.ps1        # Email sender (with backdated timestamps)
│   ├── New-DemoCalendarEvents.ps1 # Calendar event creator
│   ├── Upload-DemoFiles.ps1       # OneDrive file uploader
│   ├── Send-DemoChats.ps1         # Teams 1:1 chat sender
│   └── Initialize-DemoSharePoint.ps1  # SharePoint site + doc library
└── data/
    ├── emails.json                # Email definitions (11 emails)
    ├── calendar-events.json       # Calendar events (24 events)
    ├── files.json                 # OneDrive file manifest (5 files)
    ├── chats.json                 # Teams chat messages (12 messages)
    ├── sharepoint-files.json      # SharePoint document manifest (5 files)
    └── files/                     # Local files to upload
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
.\Load-DemoData.ps1 -WhatIf   # Preview mode
```

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
| Chats | AppOnly (client creds) | `Chat.Create`, `Chat.ReadWrite.All` | Create chats and send messages as users |
| Calendar | Delegated (interactive) | `Calendars.ReadWrite` | Events on signed-in user's calendar |
| Files | Delegated (interactive) | `Files.ReadWrite.All` | Upload to signed-in user's OneDrive |
| SharePoint | Delegated (interactive) | `Sites.ReadWrite.All`, `Group.ReadWrite.All` | Create team site + upload docs |

## Email Backdating

Emails include `dayOffset` and `time` fields to create realistic arrival timestamps. The module places messages directly into the recipient's mailbox with the correct `receivedDateTime`, making the inbox look natural for demos. Requires `Mail.ReadWrite` (Application) permission — falls back to standard send if unavailable.

## Demo Scenarios

| # | Scenario | Data Sources | Description |
|---|----------|-------------|-------------|
| 1 | **Catch Up** | Calendar | Busy week — meetings, standups, reviews |
| 2 | **Meeting Prep** | Emails + Calendar + Files + Chats | Adatum Corp customer meeting prep |
| 3 | **Research** | Emails | Adatum Corp competitive research requests |
| 4 | **Launch Plan** | Emails + Calendar | ProLine X product launch planning |
| 5 | **Write/Draft** | Emails + Files + Chats | Finish the ProLine X exec summary draft |

## Customizing Data

- **Add emails**: Edit `data/emails.json`. Set `dayOffset`/`time` for realistic timestamps.
- **Change demo week**: Update `config.json > demo > weekStart` (Monday of demo week).
- **Add files**: Place in `data/files/` and add entry to `data/files.json`.
- **Add SharePoint docs**: Add entry to `data/sharepoint-files.json` (reuses files from `data/files/`).
- **Add chats**: Add entries to `data/chats.json` with `from`, `to`, `topic`, `message`.
- **New tenant**: Copy folder, update `config.json`, run `Setup-AppRegistration.ps1`.

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
