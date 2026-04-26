<#
.SYNOPSIS
    Deploys custom Cowork skills to a user's OneDrive.
.DESCRIPTION
    Uploads SKILL.md files to /Documents/Cowork/skills/{name}/SKILL.md
    in the specified user's OneDrive. Cowork auto-discovers skills in this
    folder at the start of each conversation.
    Requires delegated auth with Files.ReadWrite.All.
#>

function Deploy-CoworkSkills {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][array]$Skills,
        [Parameter(Mandatory)][string]$DataDir
    )

    $users    = $Config.users
    $deployed = 0
    $failed   = 0

    foreach ($skill in $Skills) {
        try {
            $ownerAddr  = $users[$skill.owner].email
            $remotePath = $skill.remotePath
            $localPath  = Join-Path (Join-Path $DataDir "skills") $skill.localFile

            if (-not (Test-Path $localPath)) {
                Write-Host "  [SKIP] Skill file not found: $localPath" -ForegroundColor Yellow
                $failed++
                continue
            }

            $content   = Get-Content $localPath -Raw
            $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($content)

            $encodedPath = $remotePath -replace ' ', '%20'
            $uri = "https://graph.microsoft.com/v1.0/users/$ownerAddr/drive/root:/$encodedPath" + ":/content"

            Invoke-MgGraphRequest -Method PUT -Uri $uri `
                -ContentType "text/plain" `
                -Body $bodyBytes | Out-Null

            Write-Host "  [OK] Skill '$($skill.skillName)' -> $($users[$skill.owner].displayName)'s Cowork" -ForegroundColor Green
            $deployed++
        }
        catch {
            Write-Host "  [FAIL] Skill '$($skill.skillName)' - $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }
    }

    Write-Host "[SKILLS] $deployed deployed, $failed failed." -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Yellow' })
}
