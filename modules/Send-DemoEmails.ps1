<#
.SYNOPSIS
    Sends demo emails using Graph API with application permissions (send-as any user).
.DESCRIPTION
    Reads emails from data/emails.json and sends each one via /users/{from}/sendMail.
    Requires AppOnly auth with Mail.Send application permission.

    If emails have a "dayOffset" and "time" field, the email is created as a draft
    with a backdated receivedDateTime, then moved to the inbox — giving a realistic
    spread of emails across the demo week. Otherwise sends normally (arrives "now").
#>

function Send-DemoEmails {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][array]$Emails
    )

    $users = $Config.users
    $weekStart = if ($Config.demo.weekStart) { [datetime]::Parse($Config.demo.weekStart) } else { $null }
    $sent = 0
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
                    isRead         = $false
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
                    Write-Host "  [OK] $dayName $($email.time) - '$($email.subject)' from $($users[$email.from].displayName)" -ForegroundColor Green
                    $sent++
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
            $sent++
        }
        catch {
            Write-Host "  [FAIL] '$($email.subject)' - $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }
    }

    Write-Host "[EMAILS] $sent sent, $failed failed." -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Yellow' })
}
