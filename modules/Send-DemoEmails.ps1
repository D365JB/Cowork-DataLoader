<#
.SYNOPSIS
    Sends demo emails using Graph API with application permissions (send-as any user).
.DESCRIPTION
    Reads emails from data/emails.json and sends each one via /users/{from}/sendMail.
    Requires AppOnly auth with Mail.Send application permission.

    If emails have a "dayOffset" and "time" field, the email is created as a draft
    with a backdated receivedDateTime, then moved to the inbox — giving a realistic
    spread of emails across the demo week. Otherwise sends normally (arrives "now").

    On re-run: searches recipient's mailbox by subject. If a matching email exists,
    deletes it before creating the new one (upsert behavior, avoids duplicates).
#>

function Send-DemoEmails {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][array]$Emails
    )

    $users = $Config.users
    $weekStart = if ($Config.demo.weekStart) { [datetime]::Parse($Config.demo.weekStart) } else { $null }
    $sent = 0
    $updated = 0
    $failed = 0

    foreach ($email in $Emails) {
        try {
            $fromAddr = $users[$email.from].email
            $toAddr   = $users[$email.to].email

            $toRecipients = @(@{ emailAddress = @{ address = $toAddr } })

            $ccRecipients = @()
            if ($email.cc) {
                foreach ($ccKey in $email.cc) {
                    $ccRecipients += @{ emailAddress = @{ address = $users[$ccKey].email } }
                }
            }

            # ── Check for existing email by subject in recipient's mailbox ──
            $wasUpdate = $false
            $subjectEscaped = $email.subject -replace "'", "''"
            try {
                $searchUri = "https://graph.microsoft.com/v1.0/users/$toAddr/messages?`$filter=subject eq '$subjectEscaped'&`$select=id&`$top=10"
                $existing = Invoke-MgGraphRequest -Method GET -Uri $searchUri
                if ($existing.value.Count -gt 0) {
                    foreach ($oldMsg in $existing.value) {
                        Invoke-MgGraphRequest -Method DELETE `
                            -Uri "https://graph.microsoft.com/v1.0/users/$toAddr/messages/$($oldMsg.id)" | Out-Null
                    }
                    $wasUpdate = $true
                }
            }
            catch {
                # If search fails, proceed with creation anyway
            }

            # If email has timing info, create as draft with backdated timestamp
            if ($null -ne $email.dayOffset -and $null -ne $email.time -and $null -ne $weekStart) {
                $emailDate = $weekStart.AddDays([int]$email.dayOffset)
                $timeParts = $email.time -split ':'
                $emailDateTime = $emailDate.AddHours([int]$timeParts[0]).AddMinutes([int]$timeParts[1])
                $isoDate = $emailDateTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

                # Create the message as a draft in the recipient's inbox with a backdated receivedDateTime
                $draftMsg = @{
                    subject        = $email.subject
                    body           = @{ contentType = "HTML"; content = $email.body }
                    from           = @{ emailAddress = @{ address = $fromAddr; name = $users[$email.from].displayName } }
                    toRecipients   = $toRecipients
                    isRead         = if ($null -ne $email.isRead) { $email.isRead } else { $false }
                    receivedDateTime = $isoDate
                    sentDateTime     = $isoDate
                }
                if ($ccRecipients.Count -gt 0) { $draftMsg.ccRecipients = $ccRecipients }

                # Create directly in the recipient's mailbox (requires Mail.ReadWrite app permission)
                # Fall back to sendMail if this fails
                try {
                    Invoke-MgGraphRequest -Method POST `
                        -Uri "https://graph.microsoft.com/v1.0/users/$toAddr/messages" `
                        -Body $draftMsg | Out-Null

                    $dayName = $emailDate.ToString("ddd")
                    $status = if ($wasUpdate) { "UPD" } else { "OK" }
                    Write-Host "  [$status] $dayName $($email.time) - '$($email.subject)' from $($users[$email.from].displayName)" -ForegroundColor Green
                    if ($wasUpdate) { $updated++ } else { $sent++ }
                    continue
                }
                catch {
                    # Fall through to sendMail
                    Write-Host "  [INFO] Backdating not available, sending normally..." -ForegroundColor DarkGray
                }
            }

            # Standard send (arrives with current timestamp)
            $msg = @{
                subject      = $email.subject
                body         = @{ contentType = "HTML"; content = $email.body }
                toRecipients = $toRecipients
            }
            if ($ccRecipients.Count -gt 0) { $msg.ccRecipients = $ccRecipients }

            $params = @{ message = $msg; saveToSentItems = $true }

            Invoke-MgGraphRequest -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/users/$fromAddr/sendMail" `
                -Body $params

            Write-Host "  [OK] '$($email.subject)' from $($users[$email.from].displayName)" -ForegroundColor Green
            if ($wasUpdate) { $updated++ } else { $sent++ }
        }
        catch {
            Write-Host "  [FAIL] '$($email.subject)' - $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }
    }

    $statusColor = if ($failed -eq 0) { 'Green' } else { 'Yellow' }
    Write-Host "[EMAILS] $sent new, $updated updated, $failed failed." -ForegroundColor $statusColor
}
