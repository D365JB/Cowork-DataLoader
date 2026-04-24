<#
.SYNOPSIS
    Creates a SharePoint site and uploads shared documents for the demo.
.DESCRIPTION
    Creates a SharePoint team site (if it doesn't exist) and uploads files
    from data/sharepoint-files.json. This gives Cowork access to shared
    team documents beyond individual OneDrive files.

    Requires delegated auth with Sites.ReadWrite.All and Group.ReadWrite.All.
#>

function Initialize-DemoSharePoint {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][array]$SharePointFiles,
        [Parameter(Mandatory)][string]$DataDir
    )

    $domain = $Config.tenant.domain
    $spDomain = $domain -replace '\.OnMicrosoft\.com$', '.sharepoint.com'
    $siteName = "ApexSalesTeam"
    $siteDisplayName = "Apex Sales Team"

    # Step 1: Check if site exists, create if not
    Write-Host "  Checking for SharePoint site '$siteName'..." -ForegroundColor DarkGray
    $siteId = $null

    try {
        $siteUri = "https://graph.microsoft.com/v1.0/sites/${spDomain}:/sites/${siteName}"
        $site = Invoke-MgGraphRequest -Method GET -Uri $siteUri -ErrorAction Stop
        $siteId = $site.id
        Write-Host "  [OK] Site exists: $($site.webUrl)" -ForegroundColor Green
    }
    catch {
        # Site doesn't exist - create it via a group (Teams-connected site)
        Write-Host "  Creating new team site '$siteDisplayName'..." -ForegroundColor Cyan
        try {
            $groupBody = @{
                displayName     = $siteDisplayName
                mailNickname    = $siteName
                description     = "Apex Manufacturing Sales Team - shared documents and resources"
                groupTypes      = @("Unified")
                mailEnabled     = $true
                securityEnabled = $false
                visibility      = "Private"
            }

            # Add members
            $memberIds = @()
            foreach ($userKey in $Config.users.Keys) {
                try {
                    $userEmail = $Config.users[$userKey].email
                    $user = Invoke-MgGraphRequest -Method GET `
                        -Uri "https://graph.microsoft.com/v1.0/users/$userEmail`?`$select=id"
                    $memberIds += "https://graph.microsoft.com/v1.0/directoryObjects/$($user.id)"
                } catch {
                    Write-Host "    [WARN] Could not resolve user $userKey" -ForegroundColor Yellow
                }
            }

            if ($memberIds.Count -gt 0) {
                $groupBody["members@odata.bind"] = $memberIds
                $groupBody["owners@odata.bind"]  = @($memberIds[0])
            }

            $group = Invoke-MgGraphRequest -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/groups" -Body $groupBody

            Write-Host "  [OK] Group/site created. Waiting for site provisioning..." -ForegroundColor Green

            # Wait for SharePoint site to be provisioned (can take 10-30 seconds)
            $retries = 0
            $maxRetries = 12
            while ($retries -lt $maxRetries) {
                Start-Sleep -Seconds 5
                try {
                    $site = Invoke-MgGraphRequest -Method GET `
                        -Uri "https://graph.microsoft.com/v1.0/groups/$($group.id)/sites/root" -ErrorAction Stop
                    $siteId = $site.id
                    Write-Host "  [OK] Site provisioned: $($site.webUrl)" -ForegroundColor Green
                    break
                } catch {
                    $retries++
                    if ($retries -lt $maxRetries) {
                        Write-Host "    Waiting for site provisioning... ($retries/$maxRetries)" -ForegroundColor DarkGray
                    }
                }
            }

            if (-not $siteId) {
                Write-Host "  [FAIL] Site provisioning timed out. Files will be skipped." -ForegroundColor Red
                return
            }
        }
        catch {
            Write-Host "  [FAIL] Could not create site: $($_.Exception.Message)" -ForegroundColor Red
            return
        }
    }

    # Step 2: Upload files to the site's default document library
    $uploaded = 0
    $failed = 0

    foreach ($spFile in $SharePointFiles) {
        try {
            $remotePath = $spFile.remotePath

            if ($spFile.sourceType -eq "file") {
                $localPath = Join-Path (Join-Path $DataDir "files") $spFile.localFile
                if (-not (Test-Path $localPath)) {
                    Write-Host "  [SKIP] Local file not found: $localPath" -ForegroundColor Yellow
                    $failed++
                    continue
                }
                $content = Get-Content $localPath -Raw
            } elseif ($spFile.sourceType -eq "inline") {
                $content = $spFile.content
            } else {
                Write-Host "  [SKIP] Unknown sourceType: $($spFile.sourceType)" -ForegroundColor Yellow
                $failed++
                continue
            }

            $encodedPath = $remotePath -replace ' ', '%20'
            $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/root:/$encodedPath" + ":/content"

            Invoke-MgGraphRequest -Method PUT -Uri $uri `
                -ContentType "text/plain" `
                -Body ([System.Text.Encoding]::UTF8.GetBytes($content)) | Out-Null

            Write-Host "  [OK] $remotePath -> $siteDisplayName" -ForegroundColor Green
            $uploaded++
        }
        catch {
            Write-Host "  [FAIL] $remotePath - $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }
    }

    Write-Host "[SHAREPOINT] $uploaded uploaded, $failed failed." -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Yellow' })
}
