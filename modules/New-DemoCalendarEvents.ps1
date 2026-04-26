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
    $existingEvents = @{}
    try {
        $calUri = "https://graph.microsoft.com/v1.0/users/$ownerAddr/calendarView?startDateTime=$startIso&endDateTime=$endIso&`$select=id,subject&`$top=100"
        $result = Invoke-MgGraphRequest -Method GET -Uri $calUri
        foreach ($ev in $result.value) {
            # Store by subject (first match wins for update)
            if (-not $existingEvents[$ev.subject]) {
                $existingEvents[$ev.subject] = $ev.id
            }
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
            $existingId = $existingEvents[$evt.subject]
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
