<#
.SYNOPSIS
    One-time setup: Creates an Entra ID app registration with application
    permissions for sending emails, chats, and channel messages with correct
    sender attribution using the Teams Migration API.
.DESCRIPTION
    Run this once per tenant. It will:
    1. Create an app registration ("Cowork Demo Mail Sender")
    2. Create a client secret
    3. Create a service principal
    4. Grant application permissions with admin consent:
       - Mail.Send              (send emails as any user)
       - Mail.ReadWrite         (backdate emails into mailboxes)
       - Chat.Create            (create Teams chats)
       - Chat.ReadWrite.All     (send Teams chat messages)
       - Teamwork.Migrate.All   (import chats/channels with correct senders)
       - User.Read.All          (resolve user IDs)
       - Group.Read.All         (find Teams by name)
       - Channel.Delete.All     (delete old channels before migration reimport)

    After running, paste the clientId and clientSecret into config.json.

.PARAMETER TenantDomain
    The tenant domain (e.g., contoso.onmicrosoft.com). Uses config.json default if omitted.

.PARAMETER AppName
    Display name for the app registration. Default: "Cowork Demo Mail Sender"

.EXAMPLE
    .\Setup-AppRegistration.ps1
    # Uses tenant from config.json

.EXAMPLE
    .\Setup-AppRegistration.ps1 -TenantDomain "mytenant.onmicrosoft.com"
#>

param(
    [string]$TenantDomain,
    [string]$AppName = "Cowork Demo Mail Sender"
)

$ErrorActionPreference = "Stop"

# Load config for defaults
$configPath = Join-Path $PSScriptRoot "config.json"
if (-not $TenantDomain -and (Test-Path $configPath)) {
    $cfg = Get-Content $configPath -Raw | ConvertFrom-Json
    $TenantDomain = $cfg.tenant.domain
}

if (-not $TenantDomain) {
    Write-Host "[ERROR] No tenant domain specified and config.json not found." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "=== Cowork Demo - App Registration Setup ===" -ForegroundColor Cyan
Write-Host "  Tenant: $TenantDomain" -ForegroundColor White
Write-Host "  App:    $AppName" -ForegroundColor White
Write-Host ""

# Ensure module
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Host "Installing Microsoft.Graph.Authentication..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
}

# Connect with delegated permissions to manage apps
$env:MSAL_ENABLE_WAM = "0"
Write-Host "Connecting to Graph (delegated) - sign in as a Global Admin..." -ForegroundColor Yellow
Connect-MgGraph -Scopes "Application.ReadWrite.All","AppRoleAssignment.ReadWrite.All" `
    -TenantId $TenantDomain -NoWelcome

$ctx = Get-MgContext
if (-not $ctx) {
    Write-Host "[ERROR] Failed to connect." -ForegroundColor Red
    exit 1
}
Write-Host "[OK] Connected as $($ctx.Account)" -ForegroundColor Green
Write-Host ""

# Step 1: Create app registration
Write-Host "Step 1: Creating app registration..." -ForegroundColor Cyan
$appBody = @{
    displayName    = $AppName
    signInAudience = "AzureADMyOrg"
}
$app = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/applications" -Body $appBody
$appId    = $app.appId
$objectId = $app.id
Write-Host "  App registered: $AppName" -ForegroundColor Green
Write-Host "  Client ID: $appId" -ForegroundColor White

# Step 2: Create client secret
Write-Host ""
Write-Host "Step 2: Creating client secret..." -ForegroundColor Cyan
$secretBody = @{
    passwordCredential = @{
        displayName = "demo-secret"
        endDateTime = ([datetime]::UtcNow.AddYears(1)).ToString("yyyy-MM-ddTHH:mm:ssZ")
    }
}
$secret = Invoke-MgGraphRequest -Method POST `
    -Uri "https://graph.microsoft.com/v1.0/applications/$objectId/addPassword" -Body $secretBody
$clientSecret = $secret.secretText
Write-Host "  Secret created (expires in 1 year)" -ForegroundColor Green

# Step 3: Create service principal
Write-Host ""
Write-Host "Step 3: Creating service principal..." -ForegroundColor Cyan
$spBody = @{ appId = $appId }
$sp = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/servicePrincipals" -Body $spBody
$spId = $sp.id
Write-Host "  Service Principal ID: $spId" -ForegroundColor Green

# Step 4: Grant application permissions
Write-Host ""
Write-Host "Step 4: Granting application permissions..." -ForegroundColor Cyan

$graphSp = (Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'").value[0]

$permissionNames = @(
    "Mail.Send",
    "Mail.ReadWrite",
    "Chat.Create",
    "Chat.ReadWrite.All",
    "Teamwork.Migrate.All",
    "User.Read.All",
    "Group.Read.All",
    "Channel.Delete.All"
)

foreach ($permName in $permissionNames) {
    $role = $graphSp.appRoles | Where-Object { $_.value -eq $permName }
    if (-not $role) {
        Write-Host "  $permName - NOT FOUND in Graph permissions" -ForegroundColor Red
        continue
    }
    try {
        $roleBody = @{
            principalId = $spId
            resourceId  = $graphSp.id
            appRoleId   = $role.id
        }
        Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$spId/appRoleAssignments" -Body $roleBody | Out-Null
        Write-Host "  $permName - granted" -ForegroundColor Green
    } catch {
        if ($_.Exception.Message -match "already exists") {
            Write-Host "  $permName - already granted" -ForegroundColor DarkGray
        } else {
            Write-Host "  $permName - FAILED: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# Output
Write-Host ""
Write-Host "══════════════════════════════════════════════════" -ForegroundColor Magenta
Write-Host "  UPDATE config.json with these values:" -ForegroundColor Magenta
Write-Host "══════════════════════════════════════════════════" -ForegroundColor Magenta
Write-Host ""
Write-Host "  clientId:     $appId" -ForegroundColor Yellow
Write-Host "  clientSecret: $clientSecret" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Or run this to update config.json automatically:" -ForegroundColor DarkGray
Write-Host ""

# Auto-update config.json
$updateChoice = Read-Host "Update config.json now? (y/n)"
if ($updateChoice -eq 'y') {
    $cfg = Get-Content $configPath -Raw | ConvertFrom-Json
    $cfg.appRegistration.clientId     = $appId
    $cfg.appRegistration.clientSecret = $clientSecret
    $cfg | ConvertTo-Json -Depth 10 | Set-Content $configPath -Encoding UTF8
    Write-Host "  config.json updated!" -ForegroundColor Green
}

Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
Write-Host ""
Write-Host "Setup complete. You can now run Load-DemoData.ps1" -ForegroundColor Green
