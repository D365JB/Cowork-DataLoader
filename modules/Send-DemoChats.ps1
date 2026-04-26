<#
.SYNOPSIS
    Creates Teams 1:1 chat messages between demo users with correct sender attribution.
.DESCRIPTION
    Uses the Teams Migration API (Teamwork.Migrate.All) to import chat messages
    with the correct "from" user on each message. This ensures messages appear
    as sent by the actual demo user, not the admin.

    Flow per conversation:
    1. Create 1:1 chat in migration mode (@microsoft.graph.chatCreationMode)
    2. Import messages with from user + createdDateTime
    3. Complete migration (POST /beta/chats/{id}/completeMigration)

    If a chat already exists between the two users (from a prior run), it checks
    for existing messages and skips if found. Chat messages cannot be updated or
    deleted via Graph API.

    Requires app-only auth with Chat.Create, Chat.ReadWrite.All, and
    Teamwork.Migrate.All application permissions.

    User AAD object IDs can be set in config.json (users.*.aadObjectId) to avoid
    needing User.Read.All permission. If not set, falls back to Graph lookup.
#>

function Send-DemoChats {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][array]$Chats
    )

    $users = $Config.users
    $weekStart = if ($Config.demo.weekStart) { [datetime]::Parse($Config.demo.weekStart) } else { [datetime]::Now }
    $sent = 0
    $skipped = 0
    $failed = 0

    # ── Build user ID cache from config (aadObjectId) or Graph lookup ────────
    $userIdCache = @{}
    foreach ($userKey in $users.Keys) {
        $user = $users[$userKey]
        if ($user.aadObjectId) {
            $userIdCache[$userKey] = @{ id = $user.aadObjectId; displayName = $user.displayName }
        } else {
            try {
                $u = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/users/$($user.email)`?`$select=id,displayName"
                $userIdCache[$userKey] = @{ id = $u.id; displayName = $u.displayName }
            }
            catch {
                Write-Host "  [WARN] Could not resolve user $($user.email) : $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }

    # Group messages by conversation (from+to pair)
    $conversations = @{}
    foreach ($chat in $Chats) {
        $key = (@($chat.from, $chat.to) | Sort-Object) -join '+'
        if (-not $conversations[$key]) { $conversations[$key] = @() }
        $conversations[$key] += $chat
    }

    foreach ($convKey in $conversations.Keys) {
        $msgs = $conversations[$convKey]
        $participants = $convKey -split '\+'
        $user1 = $users[$participants[0]].email
        $user2 = $users[$participants[1]].email
        $fromName1 = $users[$participants[0]].displayName
        $fromName2 = $users[$participants[1]].displayName

        Write-Host "  [CHAT] $fromName1 <-> $fromName2" -ForegroundColor Cyan

        try {
            $migrationStart = $weekStart.AddDays(-7).ToString("yyyy-MM-ddTHH:mm:ssZ")
            $memberList = @(
                @{
                    "@odata.type"     = "#microsoft.graph.aadUserConversationMember"
                    roles             = @("owner")
                    "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$user1')"
                },
                @{
                    "@odata.type"     = "#microsoft.graph.aadUserConversationMember"
                    roles             = @("owner")
                    "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$user2')"
                }
            )

            $chatId = $null
            $inMigrationMode = $false

            # ── Try creating chat in migration mode ──────────────────────────
            try {
                $migrationChatBody = @{
                    chatType                               = "oneOnOne"
                    createdDateTime                        = $migrationStart
                    "@microsoft.graph.chatCreationMode"    = "migration"
                    members                                = $memberList
                }
                $chatResult = Invoke-MgGraphRequest -Method POST `
                    -Uri "https://graph.microsoft.com/v1.0/chats" -Body $migrationChatBody
                $chatId = $chatResult["id"]
                $inMigrationMode = $true
            }
            catch {
                # Migration-mode creation failed (chat may already exist)
                # Fall back to normal create-or-get
                $normalChatBody = @{
                    chatType = "oneOnOne"
                    members  = $memberList
                }
                $chatResult = Invoke-MgGraphRequest -Method POST `
                    -Uri "https://graph.microsoft.com/v1.0/chats" -Body $normalChatBody
                $chatId = $chatResult["id"]
            }

            # ── If not in migration mode, try startMigration on existing chat ──
            if (-not $inMigrationMode) {
                try {
                    $startBody = @{ conversationCreationDateTime = $migrationStart }
                    Invoke-MgGraphRequest -Method POST `
                        -Uri "https://graph.microsoft.com/beta/chats/$chatId/startMigration" `
                        -Body $startBody | Out-Null
                    $inMigrationMode = $true
                    Write-Host "    [OK] Entered migration mode on existing chat" -ForegroundColor Green
                }
                catch {
                    # Check if chat already has messages (the usual blocker)
                    try {
                        $chatMsgs = Invoke-MgGraphRequest -Method GET `
                            -Uri "https://graph.microsoft.com/v1.0/chats/$chatId/messages?`$top=5"
                        $realMsgs = @($chatMsgs.value | Where-Object { $_.messageType -eq 'message' })
                    } catch { $realMsgs = @() }

                    if ($realMsgs.Count -gt 0) {
                        Write-Host "    [SKIP] Chat has $($realMsgs.Count)+ existing messages from prior runs." -ForegroundColor DarkGray
                        Write-Host "    Delete the chat in Teams client and re-run to fix sender attribution." -ForegroundColor DarkGray
                    } else {
                        Write-Host "    [WARN] Cannot enter migration mode: $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                    $skipped += $msgs.Count
                    continue
                }
            }

            # ── Import messages with correct from user ───────────────────────
            $msgIndex = 0
            $importedCount = 0
            foreach ($msg in $msgs) {
                $senderKey = $msg.from
                $senderInfo = $userIdCache[$senderKey]
                if (-not $senderInfo) {
                    Write-Host "    [WARN] Unknown sender '$senderKey' - skipping message" -ForegroundColor Yellow
                    $failed++
                    continue
                }

                # Each message needs a unique createdDateTime
                $msgTime = $weekStart.AddDays(-1).AddHours(9).AddMinutes($msgIndex * 2).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")

                $msgBody = @{
                    createdDateTime = $msgTime
                    from = @{
                        user = @{
                            id              = $senderInfo.id
                            displayName     = $senderInfo.displayName
                            userIdentityType = "aadUser"
                        }
                    }
                    body = @{
                        contentType = "html"
                        content     = $msg.message
                    }
                }

                Invoke-MgGraphRequest -Method POST `
                    -Uri "https://graph.microsoft.com/v1.0/chats/$chatId/messages" `
                    -Body $msgBody | Out-Null

                $preview = ($msg.message -replace '<[^>]+>','')
                if ($preview.Length -gt 60) { $preview = $preview.Substring(0, 60) }
                Write-Host "    [OK] $($senderInfo.displayName): $preview" -ForegroundColor Green
                $sent++
                $importedCount++
                $msgIndex++
            }

            # ── Complete migration ───────────────────────────────────────────
            if ($importedCount -gt 0) {
                Invoke-MgGraphRequest -Method POST `
                    -Uri "https://graph.microsoft.com/beta/chats/$chatId/completeMigration" | Out-Null
            }
        }
        catch {
            Write-Host "    [FAIL] Chat error: $($_.Exception.Message)" -ForegroundColor Red
            $failed += $msgs.Count
        }
    }

    Write-Host "[CHATS] $sent new, $skipped skipped (existing), $failed failed." -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Yellow' })
    if ($failed -gt 0) {
        Write-Host "  TIP: Ensure app has Chat.Create, Chat.ReadWrite.All, and Teamwork.Migrate.All application permissions." -ForegroundColor DarkGray
    }
}
