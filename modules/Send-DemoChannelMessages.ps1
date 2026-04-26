<#
.SYNOPSIS
    Posts messages to Teams channels for the demo team.
.DESCRIPTION
    Creates a Team (or uses existing), creates channels, and posts messages.
    Uses delegated auth - the signed-in user posts all messages.
    Requires delegated auth with Group.ReadWrite.All, Channel.Create,
    ChannelMessage.Send scopes.
#>

function Send-DemoChannelMessages {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][array]$ChannelMessages,
        [string]$TeamDisplayName = "Apex Sales Team"
    )

    $users     = $Config.users
    $weekStart = if ($Config.demo.weekStart) { [datetime]::Parse($Config.demo.weekStart) } else { [datetime]::Now }
    $sent   = 0
    $failed = 0

    # ── Find or create the Team ──────────────────────────────────────────────
    Write-Host "  Looking for team '$TeamDisplayName'..." -ForegroundColor DarkGray

    $encodedName = [System.Uri]::EscapeDataString($TeamDisplayName)
    $teamSearch = Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/me/joinedTeams?`$filter=displayName eq '$encodedName'"

    $teamId = $null
    if ($teamSearch.value.Count -gt 0) {
        $teamId = $teamSearch.value[0].id
        Write-Host "  [OK] Found existing team: $teamId" -ForegroundColor Green
    }
    else {
        Write-Host "  Creating team '$TeamDisplayName'..." -ForegroundColor Yellow

        # Create a group first, then team-ify it
        $groupBody = @{
            displayName     = $TeamDisplayName
            description     = "Apex Manufacturing sales coordination team"
            mailEnabled     = $true
            mailNickname    = ($TeamDisplayName -replace '[^a-zA-Z0-9]', '')
            securityEnabled = $false
            groupTypes      = @("Unified")
            visibility      = "Private"
        }

        $group = Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/groups" -Body $groupBody
        $groupId = $group.id
        Write-Host "  [OK] Created group: $groupId" -ForegroundColor Green

        # Add team members
        foreach ($userKey in $users.Keys) {
            $userEmail = $users[$userKey].email
            try {
                $user = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/users/$userEmail"
                $memberBody = @{
                    "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.id)"
                }
                Invoke-MgGraphRequest -Method POST `
                    -Uri "https://graph.microsoft.com/v1.0/groups/$groupId/members/`$ref" `
                    -Body $memberBody | Out-Null
                Write-Host "    Added member: $($users[$userKey].displayName)" -ForegroundColor DarkGray
            }
            catch {
                if ($_.Exception.Message -match "already exist") {
                    Write-Host "    Member already exists: $($users[$userKey].displayName)" -ForegroundColor DarkGray
                }
                else {
                    Write-Host "    [WARN] Could not add $($users[$userKey].displayName): $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        }

        # Wait for group provisioning then create team
        Write-Host "  Waiting for group provisioning..." -ForegroundColor DarkGray
        $retries = 0
        $teamCreated = $false
        while (-not $teamCreated -and $retries -lt 10) {
            try {
                $teamBody = @{
                    memberSettings    = @{ allowCreateUpdateChannels = $true }
                    messagingSettings = @{ allowUserEditMessages = $true; allowUserDeleteMessages = $true }
                }
                Invoke-MgGraphRequest -Method PUT `
                    -Uri "https://graph.microsoft.com/v1.0/groups/$groupId/team" `
                    -Body $teamBody | Out-Null
                $teamCreated = $true
                $teamId = $groupId
                Write-Host "  [OK] Team created" -ForegroundColor Green
            }
            catch {
                $retries++
                Write-Host "    Waiting for provisioning (attempt $retries/10)..." -ForegroundColor DarkGray
                Start-Sleep -Seconds 5
            }
        }

        if (-not $teamCreated) {
            Write-Host "  [FAIL] Could not create team after 10 retries" -ForegroundColor Red
            return
        }
    }

    # ── Get or create channels ───────────────────────────────────────────────
    $channelNames = $ChannelMessages | Select-Object -ExpandProperty channel -Unique
    $channelMap = @{}

    $existingChannels = Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/channels"

    foreach ($ch in $existingChannels.value) {
        $channelMap[$ch.displayName] = $ch.id
    }

    foreach ($chName in $channelNames) {
        if (-not $channelMap[$chName]) {
            if ($chName -eq "General") {
                Write-Host "  [SKIP] General channel already exists" -ForegroundColor DarkGray
                continue
            }
            try {
                $chBody = @{
                    displayName = $chName
                    description = "$chName discussion channel"
                }
                $newCh = Invoke-MgGraphRequest -Method POST `
                    -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/channels" -Body $chBody
                $channelMap[$chName] = $newCh.id
                Write-Host "  [OK] Created channel: $chName" -ForegroundColor Green
            }
            catch {
                if ($_.Exception.Message -match "NameAlreadyExists") {
                    # Re-fetch channels
                    $existingChannels = Invoke-MgGraphRequest -Method GET `
                        -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/channels"
                    foreach ($ch in $existingChannels.value) {
                        if ($ch.displayName -eq $chName) { $channelMap[$chName] = $ch.id }
                    }
                    Write-Host "  [OK] Channel already exists: $chName" -ForegroundColor DarkGray
                }
                else {
                    Write-Host "  [FAIL] Could not create channel '$chName': $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        }
        else {
            Write-Host "  [OK] Channel exists: $chName" -ForegroundColor DarkGray
        }
    }

    # ── Post messages ────────────────────────────────────────────────────────
    foreach ($msg in ($ChannelMessages | Sort-Object { $_.dayOffset }, { $_.time })) {
        $chId = $channelMap[$msg.channel]
        if (-not $chId) {
            Write-Host "  [SKIP] No channel ID for '$($msg.channel)'" -ForegroundColor Yellow
            $failed++
            continue
        }

        try {
            $msgBody = @{
                body = @{
                    contentType = "html"
                    content     = $msg.message
                }
            }

            Invoke-MgGraphRequest -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/channels/$chId/messages" `
                -Body $msgBody | Out-Null

            $fromName = $users[$msg.from].displayName
            $preview = ($msg.message -replace '<[^>]+>','').Substring(0, [Math]::Min(60, ($msg.message -replace '<[^>]+>','').Length))
            Write-Host "  [OK] #$($msg.channel) - $fromName : $preview..." -ForegroundColor Green
            $sent++
        }
        catch {
            Write-Host "  [FAIL] #$($msg.channel) - $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }
    }

    Write-Host "[CHANNELS] $sent messages posted, $failed failed." -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Yellow' })
}
