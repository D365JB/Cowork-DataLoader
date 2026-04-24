<#
.SYNOPSIS
    Creates Teams 1:1 chat messages between demo users.
.DESCRIPTION
    Uses Graph API to create 1:1 chats and send messages via delegated auth.
    The signed-in user must be a participant in each chat.
    Messages are sent as the signed-in user (sender attribution in the JSON
    is used for display purposes in the loading output only).

    Requires delegated auth with Chat.ReadWrite permission.
#>

function Send-DemoChats {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][array]$Chats
    )

    $users = $Config.users
    $sent = 0
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

        try {
            # Create or get existing 1:1 chat
            $chatBody = @{
                chatType = "oneOnOne"
                members  = @(
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
            }

            $chatResult = Invoke-MgGraphRequest -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/chats" -Body $chatBody
            $chatId = $chatResult.id

            $fromName1 = $users[$participants[0]].displayName
            $fromName2 = $users[$participants[1]].displayName
            Write-Host "  [CHAT] $fromName1 <-> $fromName2" -ForegroundColor Cyan

            # Send each message in order
            foreach ($msg in $msgs) {
                $senderAddr = $users[$msg.from].email

                $msgBody = @{
                    body = @{
                        contentType = "html"
                        content     = $msg.message
                    }
                }

                # Use the delegated-on-behalf endpoint with app permissions
                Invoke-MgGraphRequest -Method POST `
                    -Uri "https://graph.microsoft.com/v1.0/chats/$chatId/messages" `
                    -Body $msgBody `
                    -Headers @{ "Content-Type" = "application/json" }

                Write-Host "    [OK] $($users[$msg.from].displayName): $($msg.message -replace '<[^>]+>','' | Select-Object -First 1)" -ForegroundColor Green
                $sent++
            }
        }
        catch {
            Write-Host "    [FAIL] Chat error: $($_.Exception.Message)" -ForegroundColor Red
            $failed += $msgs.Count
        }
    }

    Write-Host "[CHATS] $sent messages sent, $failed failed." -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Yellow' })
    if ($failed -gt 0) {
        Write-Host "  TIP: Ensure delegated auth includes Chat.ReadWrite scope and the signed-in user is a chat participant." -ForegroundColor DarkGray
    }
}
