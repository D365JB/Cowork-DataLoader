<#
.SYNOPSIS
    Uploads demo files to a user's OneDrive via Graph API.
.DESCRIPTION
    Reads file definitions from data/files.json. Each entry specifies:
    - owner: user key from config
    - remotePath: path in OneDrive (e.g. "Documents/filename.txt")
    - sourceType: "inline" (content in JSON) or "file" (local path in data/files/)
    Supports both text and binary files (Office docs, images, etc.).
    Requires delegated auth with Files.ReadWrite.All.
#>

# MIME type lookup by file extension
$script:MimeTypes = @{
    '.txt'  = 'text/plain'
    '.md'   = 'text/plain'
    '.csv'  = 'text/csv'
    '.json' = 'application/json'
    '.xml'  = 'application/xml'
    '.html' = 'text/html'
    '.htm'  = 'text/html'
    '.docx' = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    '.doc'  = 'application/msword'
    '.xlsx' = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    '.xls'  = 'application/vnd.ms-excel'
    '.pptx' = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    '.ppt'  = 'application/vnd.ms-powerpoint'
    '.pdf'  = 'application/pdf'
    '.png'  = 'image/png'
    '.jpg'  = 'image/jpeg'
    '.jpeg' = 'image/jpeg'
    '.gif'  = 'image/gif'
    '.zip'  = 'application/zip'
}

function Get-MimeType {
    param([string]$FileName)
    $ext = [System.IO.Path]::GetExtension($FileName).ToLower()
    if ($script:MimeTypes.ContainsKey($ext)) { return $script:MimeTypes[$ext] }
    return 'application/octet-stream'
}

function Test-BinaryFile {
    param([string]$FileName)
    $textExts = @('.txt', '.md', '.csv', '.json', '.xml', '.html', '.htm')
    $ext = [System.IO.Path]::GetExtension($FileName).ToLower()
    return ($ext -notin $textExts)
}

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
                $bodyBytes   = [System.Text.Encoding]::UTF8.GetBytes($file.content)
                $contentType = 'text/plain'
            } elseif ($file.sourceType -eq "file") {
                $localPath = Join-Path (Join-Path $DataDir "files") $file.localFile
                if (-not (Test-Path $localPath)) {
                    Write-Host "  [SKIP] Local file not found: $localPath" -ForegroundColor Yellow
                    $failed++
                    continue
                }
                $contentType = Get-MimeType $file.localFile
                if (Test-BinaryFile $file.localFile) {
                    $bodyBytes = [System.IO.File]::ReadAllBytes($localPath)
                } else {
                    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes((Get-Content $localPath -Raw))
                }
            } else {
                Write-Host "  [SKIP] Unknown sourceType: $($file.sourceType)" -ForegroundColor Yellow
                $failed++
                continue
            }

            $encodedPath = $remotePath -replace ' ', '%20'
            $uri = "https://graph.microsoft.com/v1.0/users/$ownerAddr/drive/root:/$encodedPath" + ":/content"

            Invoke-MgGraphRequest -Method PUT -Uri $uri `
                -ContentType $contentType `
                -Body $bodyBytes | Out-Null

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
