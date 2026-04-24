<#
.SYNOPSIS
    Creates demo calendar events using Graph API with delegated permissions.
.DESCRIPTION
    Reads events from data/calendar-events.json and creates each one via /users/{owner}/events.
    Supports attendees, all-day events, and conflict generation.
    Uses weekStart from config.json to calculate actual dates (dayOffset 0 = Monday).
#>

function New-DemoCalendarEvents {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][array]$Events
    )

    $users     = $Config.users
    $weekStart = [datetime]::Parse($Config.demo.weekStart)
    $timeZone  = "Eastern Standard Time"
    $created   = 0
    $failed    = 0

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

            Invoke-MgGraphRequest -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/users/$ownerAddr/events" `
                -Body $body | Out-Null

            $dayName = $eventDate.ToString("ddd")
            $timeStr = if ($evt.allDay) { "(all day)" } else { "$($evt.startTime)-$($evt.endTime)" }
            Write-Host "  [OK] $dayName $timeStr - $($evt.subject)" -ForegroundColor Green
            $created++
        }
        catch {
            Write-Host "  [FAIL] $($evt.subject) - $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }
    }

    Write-Host "[CALENDAR] $created created, $failed failed." -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Yellow' })
}
