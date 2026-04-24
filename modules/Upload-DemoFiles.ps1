<#
.SYNOPSIS
    Uploads demo files to a user's OneDrive via Graph API.
.DESCRIPTION
    Reads file definitions from data/files.json. Each entry specifies:
    - owner: user key from config
    - remotePath: path in OneDrive (e.g. "Documents/filename.txt")
    - sourceType: "inline" (content in JSON) or "file" (local path in data/files/)
    Requires delegated auth with Files.ReadWrite.All.
#>

function Upload-DemoFiles {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][array]$Files,
        [Parameter(Mandatory)][string]$DataDir
    )

    $users    = $Config.users
    $uploaded = 0
    $failed   = 0

    foreach ($file in $Files) {
        try {
            $ownerAddr  = $users[$file.owner].email
            $remotePath = $file.remotePath

            if ($file.sourceType -eq "inline") {
                $content = $file.content
            } elseif ($file.sourceType -eq "file") {
                $localPath = Join-Path $DataDir "files" $file.localFile
                if (-not (Test-Path $localPath)) {
                    Write-Host "  [SKIP] Local file not found: $localPath" -ForegroundColor Yellow
                    $failed++
                    continue
                }
                $content = Get-Content $localPath -Raw
            } else {
                Write-Host "  [SKIP] Unknown sourceType: $($file.sourceType)" -ForegroundColor Yellow
                $failed++
                continue
            }

            $encodedPath = $remotePath -replace ' ', '%20'
            $uri = "https://graph.microsoft.com/v1.0/users/$ownerAddr/drive/root:/$encodedPath" + ":/content"

            Invoke-MgGraphRequest -Method PUT -Uri $uri `
                -ContentType "text/plain" `
                -Body ([System.Text.Encoding]::UTF8.GetBytes($content)) | Out-Null

            Write-Host "  [OK] $remotePath -> $($users[$file.owner].displayName)'s OneDrive" -ForegroundColor Green
            $uploaded++
        }
        catch {
            Write-Host "  [FAIL] $remotePath - $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }
    }

    Write-Host "[FILES] $uploaded uploaded, $failed failed." -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Yellow' })
}
