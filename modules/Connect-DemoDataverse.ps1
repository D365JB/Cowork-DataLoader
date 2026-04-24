<#
.SYNOPSIS
    Connects to Dataverse Web API for D365 record management.
.DESCRIPTION
    Authenticates using client credentials (app-only) to the Dataverse environment.
    Requires the app registration to have Dynamics CRM user_impersonation or
    application permission on the Dataverse environment.
#>

function Connect-DemoDataverse {
    param(
        [Parameter(Mandatory)][hashtable]$Config
    )

    $envUrl = $Config.dataverse.environmentUrl.TrimEnd('/')
    $tenantId = $Config.tenant.tenantId
    $clientId = $Config.appRegistration.clientId
    $clientSecret = $Config.appRegistration.clientSecret

    if (-not $envUrl) {
        Write-Host "[AUTH] No dataverse.environmentUrl in config.json - skipping D365." -ForegroundColor Yellow
        return $null
    }

    Write-Host "[AUTH] Connecting to Dataverse: $envUrl" -ForegroundColor Yellow

    # Get OAuth token via client credentials
    $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $clientId
        client_secret = $clientSecret
        scope         = "$envUrl/.default"
    }

    try {
        $response = Invoke-RestMethod -Method POST -Uri $tokenUrl -Body $body -ContentType "application/x-www-form-urlencoded"
        $token = $response.access_token

        if ($token) {
            Write-Host "[AUTH] Connected to Dataverse (AppOnly)" -ForegroundColor Green
            return @{
                Token   = $token
                BaseUrl = "$envUrl/api/data/v9.2"
                Headers = @{
                    "Authorization" = "Bearer $token"
                    "OData-MaxVersion" = "4.0"
                    "OData-Version" = "4.0"
                    "Accept" = "application/json"
                    "Content-Type" = "application/json; charset=utf-8"
                    "Prefer" = "return=representation"
                }
            }
        } else {
            Write-Host "[AUTH] Failed - no token returned." -ForegroundColor Red
            return $null
        }
    } catch {
        Write-Host "[AUTH] Dataverse auth failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "[AUTH] Ensure the app registration has Dynamics CRM API permission and an application user exists in the environment." -ForegroundColor Yellow
        return $null
    }
}

function Invoke-DataverseRequest {
    param(
        [Parameter(Mandatory)][hashtable]$Connection,
        [Parameter(Mandatory)][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [object]$Body = $null
    )

    $fullUri = if ($Uri.StartsWith("http")) { $Uri } else { "$($Connection.BaseUrl)/$Uri" }
    $params = @{
        Method  = $Method
        Uri     = $fullUri
        Headers = $Connection.Headers
    }

    if ($Body) {
        $params.Body = ($Body | ConvertTo-Json -Depth 10)
    }

    return Invoke-RestMethod @params
}
