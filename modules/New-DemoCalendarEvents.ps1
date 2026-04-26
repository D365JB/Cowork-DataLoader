<#
.SYNOPSIS
    Creates demo calendar events using Graph API with delegated permissions.
.DESCRIPTION
    Reads events from data/calendar-events.json and creates each one via /users/{owner}/events.
    Supports attendees, all-day events, and conflict generation.
    Uses weekStart from config.json to calculate actual dates (dayOffset 0 = Monday).

    On re-run: searches calendar for events with matching subject in the demo week.
    If found, updates (PATCH) the existing event. If not, creates a new one.
#>

function New-DemoCalendarEvents {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][array]$Events
    )

    $users     = $Config.users
    $weekStart = [datetime]::Parse($Config.demo.weekStart)
    $weekEnd   = $weekStart.AddDays(7)
    $timeZone  = "Eastern Standard Time"
    $created   = 0
    $updated   = 0
    $failed    = 0

    # ── Pre-fetch existing events for the demo week ──────────────────────────
    $ownerAddr = $users[($Events[0].owner)].email
    $startIso  = $weekStart.ToString("yyyy-MM-ddT00:00:00")
    $endIso    = $weekEnd.ToString("yyyy-MM-ddT00:00:00")
    $existingEvents = @{}      # key = "subject|startDateTime" → first event id
    $existingDupes  = @()      # duplicate event ids to clean up
    try {
        $calUri = "https://graph.microsoft.com/v1.0/users/$ownerAddr/calendarView?startDateTime=$startIso&endDateTime=$endIso&`$select=id,subject,start&`$top=200"
        $headers = @{ Prefer = "outlook.timezone=""$timeZone""" }
        $result = Invoke-MgGraphRequest -Method GET -Uri $calUri -Headers $headers
        foreach ($ev in $result.value) {
            $normDt = ([datetime]::Parse($ev.start.dateTime)).ToString("yyyy-MM-ddTHH:mm")
            $key = "$($ev.subject)|$normDt"
            if (-not $existingEvents[$key]) {
                $existingEvents[$key] = $ev.id
            }
            else {
                # Duplicate — mark for cleanup
                $existingDupes += $ev.id
            }
        }
        # Clean up duplicates from prior runs
        if ($existingDupes.Count -gt 0) {
            Write-Host "  [CLEANUP] Removing $($existingDupes.Count) duplicate events..." -ForegroundColor Yellow
            foreach ($dupeId in $existingDupes) {
                try {
                    Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/users/$ownerAddr/events/$dupeId" | Out-Null
                } catch { }
            }
            Write-Host "  [CLEANUP] Done." -ForegroundColor Green
        }
    }
    catch {
        # If pre-fetch fails, proceed with create-only mode
    }

    foreach ($evt in $Events) {
        try {
            $ownerAddr = $users[$evt.owner].email
            $eventDate = $weekStart.AddDays($evt.dayOffset)

            if ($evt.allDay -eq $true) {
                $body = @{
                    subject       = $evt.subject
                    isAllDay      = $true
                    start         = @{ dateTime = $eventDate.ToString("yyyy-MM-ddT00:00:00"); timeZone = $timeZone }
                    end           = @{ dateTime = $eventDate.AddDays(1).ToString("yyyy-MM-ddT00:00:00"); timeZone = $timeZone }
                    isReminderOn  = $false
                }
            } else {
                $startTime = $eventDate.Add([timespan]::Parse($evt.startTime))
                $endTime   = $eventDate.Add([timespan]::Parse($evt.endTime))
                $body = @{
                    subject = $evt.subject
                    start   = @{ dateTime = $startTime.ToString("yyyy-MM-ddTHH:mm:ss"); timeZone = $timeZone }
                    end     = @{ dateTime = $endTime.ToString("yyyy-MM-ddTHH:mm:ss"); timeZone = $timeZone }
                }
            }

            if ($evt.location) { $body.location = @{ displayName = $evt.location } }
            if ($evt.body)     { $body.body = @{ contentType = "HTML"; content = $evt.body } }

            if ($evt.attendees) {
                $body.attendees = @()
                foreach ($attKey in $evt.attendees) {
                    $body.attendees += @{
                        emailAddress = @{ address = $users[$attKey].email; name = $users[$attKey].displayName }
                        type         = "required"
                    }
                }
            }

            # ── Update existing or create new ────────────────────────────────
            $normDt    = ([datetime]::Parse($body.start.dateTime)).ToString("yyyy-MM-ddTHH:mm")
            $lookupKey  = "$($evt.subject)|$normDt"
            $existingId = $existingEvents[$lookupKey]
            $dayName = $eventDate.ToString("ddd")
            $timeStr = if ($evt.allDay) { "(all day)" } else { "$($evt.startTime)-$($evt.endTime)" }

            if ($existingId) {
                Invoke-MgGraphRequest -Method PATCH `
                    -Uri "https://graph.microsoft.com/v1.0/users/$ownerAddr/events/$existingId" `
                    -Body $body | Out-Null
                Write-Host "  [UPD] $dayName $timeStr - $($evt.subject)" -ForegroundColor Green
                $updated++
            }
            else {
                Invoke-MgGraphRequest -Method POST `
                    -Uri "https://graph.microsoft.com/v1.0/users/$ownerAddr/events" `
                    -Body $body | Out-Null
                Write-Host "  [OK] $dayName $timeStr - $($evt.subject)" -ForegroundColor Green
                $created++
            }
        }
        catch {
            Write-Host "  [FAIL] $($evt.subject) - $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }
    }

    Write-Host "[CALENDAR] $created new, $updated updated, $failed failed." -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Yellow' })
}
