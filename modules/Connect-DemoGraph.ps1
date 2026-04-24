<#
.SYNOPSIS
    Connects to Microsoft Graph with the appropriate auth for each data type.
.DESCRIPTION
    - AppOnly (client credentials): Used for sending mail as any user.
    - Delegated (interactive): Used for calendar events, OneDrive files, etc.
#>

function Connect-DemoGraphAppOnly {
    param(
        [Parameter(Mandatory)][hashtable]$Config
    )

    $tenantId = $Config.tenant.tenantId
    $clientId = $Config.appRegistration.clientId
    $secret   = $Config.appRegistration.clientSecret

    if ([string]::IsNullOrWhiteSpace($secret)) {
        $secureSecret = Read-Host "Enter client secret for '$($Config.appRegistration.displayName)'" -AsSecureString
    } else {
        $secureSecret = ConvertTo-SecureString $secret -AsPlainText -Force
    }

    $cred = New-Object System.Management.Automation.PSCredential($clientId, $secureSecret)

    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $cred -NoWelcome

    $ctx = Get-MgContext
    if ($ctx -and $ctx.AuthType -eq 'AppOnly') {
        Write-Host "[AUTH] Connected as app: $($ctx.AppName) (AppOnly)" -ForegroundColor Green
        return $true
    } else {
        Write-Host "[AUTH] Failed to connect with app credentials." -ForegroundColor Red
        return $false
    }
}

function Connect-DemoGraphDelegated {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string[]]$Scopes = @("Calendars.ReadWrite", "Files.ReadWrite.All", "User.Read.All")
    )

    $tenantId = $Config.tenant.tenantId

    # Check if already connected with a delegated session
    $ctx = Get-MgContext
    if ($ctx -and $ctx.Account) {
        Write-Host "[AUTH] Already connected as $($ctx.Account) (Delegated)" -ForegroundColor Green
        return $true
    }

    # Use interactive browser auth — look for browser popup (Alt+Tab if needed)
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-Host "[AUTH] Opening browser for sign-in... (Alt+Tab if you don't see it)" -ForegroundColor Yellow
    $env:MSAL_ENABLE_WAM = "0"
    try {
        Connect-MgGraph -Scopes $Scopes -TenantId "$($Config.tenant.domain)" -ErrorAction Stop -NoWelcome
    } catch {
        Write-Host "[AUTH] Browser auth error: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "[AUTH] Checking if auth succeeded anyway..." -ForegroundColor Yellow
    }

    $ctx = Get-MgContext
    if ($ctx -and $ctx.Account) {
        Write-Host "[AUTH] Connected as $($ctx.Account) (Delegated)" -ForegroundColor Green
        return $true
    } else {
        Write-Host "[AUTH] Failed to connect with delegated auth." -ForegroundColor Red
        return $false
    }
}
