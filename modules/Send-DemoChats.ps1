<#
.SYNOPSIS
    Creates Teams chat messages between demo users.
.DESCRIPTION
    Uses delegated auth to create chats and send messages. Tries 1:1 chats
    first; if the thread is soft-deleted (compliance-deleted), falls back to
    a 2-person group chat which gets a new random ID.

    Because the Teams Migration API does not support chat threads, all
    messages are sent as the signed-in admin user.

    If a chat already has messages from a prior run, it skips that conversation
    to avoid duplicates.

    Requires delegated auth with Chat.ReadWrite scope.
#>

function Send-DemoChats {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][array]$Chats
    )

    $users = $Config.users
    $sent = 0
    $skipped = 0
    $failed = 0

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
            # ── Build member list (reused for oneOnOne and group fallback) ───
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
            $usedGroupFallback = $false

            # ── Try 1:1 chat first ──────────────────────────────────────────
            $chatBody = @{ chatType = "oneOnOne"; members = $memberList }
            $chatJson = $chatBody | ConvertTo-Json -Depth 5 -Compress
            $chatResult = Invoke-MgGraphRequest -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/chats" -Body $chatJson
            $chatId = $chatResult["id"]

            # ── Check if the thread is accessible (soft-deleted threads fail here)
            $threadOk = $true
            try {
                $existing = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/chats/$chatId/messages?`$top=5"
                $realMsgs = @($existing.value | Where-Object { $_.messageType -eq 'message' })
            }
            catch {
                if ($_.Exception.Message -match 'NotFound|SoftDeleted') {
                    $threadOk = $false
                } else {
                    $realMsgs = @()
                }
            }

            # ── Fall back to 2-person group chat if 1:1 thread is broken ────
            if (-not $threadOk) {
                Write-Host "    [INFO] 1:1 thread soft-deleted, using group chat fallback" -ForegroundColor DarkYellow
                $groupBody = @{ chatType = "group"; members = $memberList }
                $groupJson = $groupBody | ConvertTo-Json -Depth 5 -Compress
                $groupResult = Invoke-MgGraphRequest -Method POST `
                    -Uri "https://graph.microsoft.com/v1.0/chats" -Body $groupJson
                $chatId = $groupResult["id"]
                $usedGroupFallback = $true

                # Check new group chat for existing messages
                try {
                    $existing = Invoke-MgGraphRequest -Method GET `
                        -Uri "https://graph.microsoft.com/v1.0/chats/$chatId/messages?`$top=5"
                    $realMsgs = @($existing.value | Where-Object { $_.messageType -eq 'message' })
                } catch { $realMsgs = @() }
            }

            if ($realMsgs.Count -gt 0) {
                Write-Host "    [SKIP] Chat has $($realMsgs.Count)+ existing messages." -ForegroundColor DarkGray
                $skipped += $msgs.Count
                continue
            }

            # ── Send messages ────────────────────────────────────────────────
            foreach ($msg in $msgs) {
                $senderName = $users[$msg.from].displayName

                $msgBody = @{
                    body = @{
                        contentType = "html"
                        content     = $msg.message
                    }
                }
                $msgJson = $msgBody | ConvertTo-Json -Depth 3 -Compress
                Invoke-MgGraphRequest -Method POST `
                    -Uri "https://graph.microsoft.com/v1.0/chats/$chatId/messages" `
                    -Body $msgJson | Out-Null

                $preview = ($msg.message -replace '<[^>]+>','')
                if ($preview.Length -gt 60) { $preview = $preview.Substring(0, 60) }
                Write-Host "    [OK] ($senderName): $preview" -ForegroundColor Green
                $sent++
            }
        }
        catch {
            Write-Host "    [FAIL] Chat error: $($_.Exception.Message)" -ForegroundColor Red
            $failed += $msgs.Count
        }
    }

    Write-Host "[CHATS] $sent sent, $skipped skipped (existing), $failed failed." -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Yellow' })
    if ($sent -gt 0) {
        Write-Host "  NOTE: All messages sent as admin user (Teams API limitation)." -ForegroundColor DarkGray
    }
}
