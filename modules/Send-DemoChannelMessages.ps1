<#
.SYNOPSIS
    Posts messages to Teams channels with correct sender attribution.
.DESCRIPTION
    Uses the Teams Migration API (Teamwork.Migrate.All) to import channel messages
    with the correct "from" user on each message.

    Flow per non-General channel:
    1. If channel exists: use startMigration to enter migration mode (wipes messages)
    2. If channel doesn't exist: create it, then startMigration
    3. Import messages with from user + createdDateTime
    4. Complete migration via completeMigration (beta)

    If channel creation fails with NameAlreadyExists (soft-deleted channel blocks
    the name), a temp-named channel is created and renamed after migration.

    The General channel cannot use startMigration (Spacetype 2).
    General channel messages are skipped with a note.

    Requires app-only auth with Teamwork.Migrate.All, Group.ReadWrite.All,
    Group.Read.All, Channel.Delete.All, and ChannelSettings.ReadWrite.All.
#>

function Send-DemoChannelMessages {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][array]$ChannelMessages
    )

    $users     = $Config.users
    $weekStart = if ($Config.demo.weekStart) { [datetime]::Parse($Config.demo.weekStart) } else { [datetime]::Now }
    $TeamDisplayName = if ($Config.demo.teamName) { $Config.demo.teamName } else { "Contoso" }
    $sent   = 0
    $skipped = 0
    $failed = 0

    # ── If demo week is in the future, shift timestamps to recent past ───────
    # Migration API rejects "Creation time is in the future" errors.
    # Shift so last message lands ~5 min before now, preserving relative order.
    $now = [datetime]::UtcNow
    $effectiveWeekStart = $weekStart
    $lastDay = ($ChannelMessages | ForEach-Object { [int]$_.dayOffset } | Measure-Object -Maximum).Maximum
    $latestMsgTime = $weekStart.AddDays($lastDay).Date.AddHours(23).AddMinutes(59)
    if ($latestMsgTime -gt $now) {
        $shift = $latestMsgTime - $now
        $shiftDays = [math]::Ceiling($shift.TotalDays) + 1
        $effectiveWeekStart = $weekStart.AddDays(-$shiftDays)
        Write-Host "  [INFO] Demo week is in the future - shifting message timestamps back $shiftDays days for migration import" -ForegroundColor DarkGray
    }

    # ── Build user ID cache from config (aadObjectId) or Graph lookup ────────
    $userIdCache = @{}
    foreach ($userKey in $users.Keys) {
        $user = $users[$userKey]
        if ($user.aadObjectId) {
            $userIdCache[$userKey] = @{ id = $user.aadObjectId; displayName = $user.displayName }
        } else {
            try {
                $u = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/users/$($user.email)?`$select=id,displayName"
                $userIdCache[$userKey] = @{ id = $u.id; displayName = $u.displayName }
            }
            catch {
                Write-Host "  [WARN] Could not resolve user $($user.email): $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }

    # ── Find team by name (app-only compatible) ──────────────────────────────
    Write-Host "  Looking for team '$TeamDisplayName'..." -ForegroundColor DarkGray

    $filter = "displayName eq '$TeamDisplayName'"
    $teamSearch = Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=$filter&`$select=id,displayName,resourceProvisioningOptions"

    $teamId = $null
    foreach ($g in $teamSearch.value) {
        if ($g.resourceProvisioningOptions -contains 'Team') {
            $teamId = $g.id
            break
        }
    }

    if (-not $teamId) {
        Write-Host "  [FAIL] Team '$TeamDisplayName' not found." -ForegroundColor Red
        Write-Host "  Create the team in Teams UI first, then re-run." -ForegroundColor Yellow
        return
    }
    Write-Host "  [OK] Found team: $teamId" -ForegroundColor Green

    # ── Gather channel names from data ───────────────────────────────────────
    $channelNames = @($ChannelMessages | ForEach-Object { $_.channel } | Select-Object -Unique)

    # ── Process each channel ─────────────────────────────────────────────────
    foreach ($chName in $channelNames) {
        $chMsgs = @($ChannelMessages | Where-Object { $_.channel -eq $chName } |
                     Sort-Object { [int]$_.dayOffset }, { $_.time })

        # ── General channel: skip (can't use startMigration) ─────────────────
        if ($chName -eq "General") {
            Write-Host "  [SKIP] #General - cannot use migration on built-in channel ($($chMsgs.Count) msgs)" -ForegroundColor Yellow
            $skipped += $chMsgs.Count
            continue
        }

        Write-Host "  [CHANNEL] $chName ($($chMsgs.Count) messages)" -ForegroundColor Cyan

        # ── Fetch existing channels ──────────────────────────────────────────
        $existingChannels = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/channels"
        $existingId = $null
        foreach ($ch in $existingChannels.value) {
            if ($ch.displayName -eq $chName) { $existingId = $ch.id; break }
        }

        $chId = $null
        $needsRename = $false

        if ($existingId) {
            # ── Reuse existing channel ───────────────────────────────────────
            $chId = $existingId
            Write-Host "    [OK] Found existing channel: $chId" -ForegroundColor Green
        } else {
            # ── Create channel ───────────────────────────────────────────────
            try {
                $chBody = @{
                    displayName    = $chName
                    description    = "$chName discussion channel"
                    membershipType = "standard"
                } | ConvertTo-Json -Compress
                $newCh = Invoke-MgGraphRequest -Method POST `
                    -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/channels" -Body $chBody
                $chId = $newCh.id
                Write-Host "    [OK] Created channel: $chId" -ForegroundColor Green
            }
            catch {
                $errMsg = $_.Exception.Message
                $errDetail = ""
                if ($_.ErrorDetails.Message) { $errDetail = $_.ErrorDetails.Message }
                if ($errMsg -match 'NameAlreadyExists|name already existed' -or $errDetail -match 'NameAlreadyExists|name already existed') {
                    # Soft-deleted channel blocks the name - use temp name
                    $tempName = "$chName $(Get-Date -Format 'MMdd-HHmm')"
                    Write-Host "    [WARN] Name blocked by soft-deleted channel, using temp name: $tempName" -ForegroundColor Yellow
                    try {
                        $chBody = @{
                            displayName    = $tempName
                            description    = "$chName discussion channel"
                            membershipType = "standard"
                        } | ConvertTo-Json -Compress
                        $newCh = Invoke-MgGraphRequest -Method POST `
                            -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/channels" -Body $chBody
                        $chId = $newCh.id
                        $needsRename = $true
                        Write-Host "    [OK] Created temp channel: $chId" -ForegroundColor Green
                    }
                    catch {
                        Write-Host "    [FAIL] Cannot create temp channel: $($_.Exception.Message)" -ForegroundColor Red
                        $failed += $chMsgs.Count
                        continue
                    }
                } else {
                    Write-Host "    [FAIL] Cannot create channel: $errMsg $errDetail" -ForegroundColor Red
                    $failed += $chMsgs.Count
                    continue
                }
            }
        }

        # ── Start migration on the channel ───────────────────────────────────
        $migrationStart = [datetime]::UtcNow.AddDays(-30).ToString("yyyy-MM-ddTHH:mm:ssZ")
        try {
            $migBody = @{ conversationCreationDateTime = $migrationStart } | ConvertTo-Json -Compress
            Invoke-MgGraphRequest -Method POST `
                -Uri "https://graph.microsoft.com/beta/teams/$teamId/channels/$chId/startMigration" `
                -Body $migBody
            Write-Host "    [OK] Channel in migration mode" -ForegroundColor Green
        }
        catch {
            Write-Host "    [FAIL] startMigration: $($_.Exception.Message)" -ForegroundColor Red
            $failed += $chMsgs.Count
            continue
        }

        # ── Import messages with correct from user ───────────────────────────
        $importedCount = 0
        foreach ($msg in $chMsgs) {
            $senderKey = $msg.from
            $senderInfo = $userIdCache[$senderKey]
            if (-not $senderInfo) {
                Write-Host "    [WARN] Unknown sender '$senderKey' - skipping" -ForegroundColor Yellow
                $failed++
                continue
            }

            $dayOffset = [int]$msg.dayOffset
            $timeParts = $msg.time -split ':'
            $dt = $effectiveWeekStart.AddDays($dayOffset).Date.AddHours([int]$timeParts[0]).AddMinutes([int]$timeParts[1])
            $msgTime = $dt.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")

            $msgBody = @{
                createdDateTime = $msgTime
                from = @{
                    user = @{
                        id               = $senderInfo.id
                        displayName      = $senderInfo.displayName
                        userIdentityType = "aadUser"
                    }
                }
                body = @{
                    contentType = "html"
                    content     = $msg.message
                }
            } | ConvertTo-Json -Depth 5 -Compress

            try {
                Invoke-MgGraphRequest -Method POST `
                    -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/channels/$chId/messages" `
                    -Body $msgBody | Out-Null

                $preview = ($msg.message -replace '<[^>]+>','')
                if ($preview.Length -gt 60) { $preview = $preview.Substring(0, 60) }
                Write-Host "    [OK] $($senderInfo.displayName): $preview" -ForegroundColor Green
                $sent++
                $importedCount++
            }
            catch {
                $msgErr = $_.Exception.Message
                $msgErrDetail = ""
                if ($_.ErrorDetails.Message) { $msgErrDetail = $_.ErrorDetails.Message }
                Write-Host "    [FAIL] $msgErr" -ForegroundColor Red
                if ($msgErrDetail) { Write-Host "           $msgErrDetail" -ForegroundColor DarkRed }
                $failed++
            }
        }

        # ── Complete migration ───────────────────────────────────────────────
        if ($importedCount -gt 0) {
            try {
                Invoke-MgGraphRequest -Method POST `
                    -Uri "https://graph.microsoft.com/beta/teams/$teamId/channels/$chId/completeMigration"
                Write-Host "    [OK] Migration completed" -ForegroundColor Green
            }
            catch {
                Write-Host "    [WARN] completeMigration: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }

        # ── Rename temp channel to target name ───────────────────────────────
        if ($needsRename -and $importedCount -gt 0) {
            # Trigger SharePoint provisioning by accessing filesFolder
            Write-Host "    Triggering SharePoint provisioning..." -ForegroundColor DarkGray
            try {
                Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/channels/$chId/filesFolder" | Out-Null
            } catch { }
            Start-Sleep -Seconds 5

            $renamed = $false
            for ($retry = 1; $retry -le 4; $retry++) {
                try {
                    $renameBody = @{ displayName = $chName } | ConvertTo-Json -Compress
                    Invoke-MgGraphRequest -Method PATCH `
                        -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/channels/$chId" `
                        -Body $renameBody
                    Write-Host "    [OK] Renamed to '$chName'" -ForegroundColor Green
                    $renamed = $true
                    break
                }
                catch {
                    $renameErr = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
                    if ($renameErr -match 'NameAlreadyExists|name already existed') {
                        Write-Host "    [WARN] Cannot rename - soft-deleted channel blocks the name '$chName'" -ForegroundColor Yellow
                        break
                    }
                    if ($retry -lt 4) {
                        Write-Host "    [RETRY] Rename attempt $retry/$([int]4) (SharePoint provisioning), waiting 15s..." -ForegroundColor DarkGray
                        Start-Sleep -Seconds 15
                    }
                }
            }
            if (-not $renamed) {
                Write-Host "    [WARN] Could not rename channel. Rename manually in Teams." -ForegroundColor Yellow
            }
        }
    }

    Write-Host "[CHANNELS] $sent imported, $skipped skipped (General), $failed failed." `
        -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Yellow' })
    if ($failed -gt 0) {
        Write-Host "  TIP: Ensure app has Teamwork.Migrate.All, Group.ReadWrite.All, Channel.Delete.All, and ChannelSettings.ReadWrite.All permissions." -ForegroundColor DarkGray
    }
}
