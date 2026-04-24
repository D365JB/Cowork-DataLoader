<#
.SYNOPSIS
    Creates D365 Sales demo records (accounts, contacts, opportunities, notes).
.DESCRIPTION
    Uses the Dataverse Web API to create CRM records that align with the
    Copilot Cowork demo story. Designed for Copilot for Sales integration
    so meeting prep, account summaries, and opportunity details surface
    alongside the M365 data.

    Requires an authenticated Dataverse connection from Connect-DemoDataverse.
#>

function Initialize-DemoD365 {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][hashtable]$Connection,
        [Parameter(Mandatory)]$D365Records
    )

    $weekStart = [datetime]::Parse($Config.demo.weekStart)
    $accountIds = @{}
    $contactIds = @{}
    $opportunityIds = @{}

    # ── Accounts ─────────────────────────────────────────────────────────────

    Write-Host "  ACCOUNTS:" -ForegroundColor White
    $created = 0; $failed = 0
    foreach ($acct in $D365Records.accounts) {
        try {
            # Check if account already exists by name
            $existing = Invoke-DataverseRequest -Connection $Connection -Method GET `
                -Uri "accounts?`$filter=name eq '$($acct.name -replace "'","''")'&`$select=accountid,name&`$top=1"

            if ($existing.value -and $existing.value.Count -gt 0) {
                $accountIds[$acct.key] = $existing.value[0].accountid
                Write-Host "    [EXISTS] $($acct.name)" -ForegroundColor DarkGray
                continue
            }

            $body = @{
                name                          = $acct.name
                industrycode                  = Get-IndustryCode $acct.industry
                revenue                       = $acct.revenue
                numberofemployees             = $acct.employees
                websiteurl                    = $acct.website
                telephone1                    = $acct.phone
                address1_line1                = $acct.address.street
                address1_city                 = $acct.address.city
                address1_stateorprovince      = $acct.address.state
                address1_postalcode           = $acct.address.zip
                address1_country              = $acct.address.country
                description                   = $acct.description
            }

            $result = Invoke-DataverseRequest -Connection $Connection -Method POST -Uri "accounts" -Body $body
            $accountIds[$acct.key] = $result.accountid
            Write-Host "    [OK] $($acct.name)" -ForegroundColor Green
            $created++
        } catch {
            Write-Host "    [FAIL] $($acct.name): $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }
    }
    Write-Host "  [ACCOUNTS] $created created, $failed failed." -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Yellow' })

    # ── Contacts ─────────────────────────────────────────────────────────────

    Write-Host ""
    Write-Host "  CONTACTS:" -ForegroundColor White
    $created = 0; $failed = 0
    foreach ($contact in $D365Records.contacts) {
        try {
            # Check if contact already exists
            $existing = Invoke-DataverseRequest -Connection $Connection -Method GET `
                -Uri "contacts?`$filter=firstname eq '$($contact.firstName)' and lastname eq '$($contact.lastName)'&`$select=contactid&`$top=1"

            if ($existing.value -and $existing.value.Count -gt 0) {
                $contactIds[$contact.key] = $existing.value[0].contactid
                Write-Host "    [EXISTS] $($contact.firstName) $($contact.lastName)" -ForegroundColor DarkGray
                continue
            }

            $body = @{
                firstname   = $contact.firstName
                lastname    = $contact.lastName
                jobtitle    = $contact.jobTitle
                emailaddress1 = $contact.email
                telephone1  = $contact.phone
                description = $contact.description
            }

            # Link to parent account if we have it
            $parentAccountId = $accountIds[$contact.account]
            if ($parentAccountId) {
                $body["parentcustomerid_account@odata.bind"] = "/accounts($parentAccountId)"
            }

            $result = Invoke-DataverseRequest -Connection $Connection -Method POST -Uri "contacts" -Body $body
            $contactIds[$contact.key] = $result.contactid
            Write-Host "    [OK] $($contact.firstName) $($contact.lastName) ($($contact.jobTitle))" -ForegroundColor Green
            $created++
        } catch {
            Write-Host "    [FAIL] $($contact.firstName) $($contact.lastName): $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }
    }
    Write-Host "  [CONTACTS] $created created, $failed failed." -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Yellow' })

    # ── Opportunities ────────────────────────────────────────────────────────

    Write-Host ""
    Write-Host "  OPPORTUNITIES:" -ForegroundColor White
    $created = 0; $failed = 0
    foreach ($opp in $D365Records.opportunities) {
        try {
            # Check if opportunity already exists
            $existing = Invoke-DataverseRequest -Connection $Connection -Method GET `
                -Uri "opportunities?`$filter=name eq '$($opp.name -replace "'","''")'&`$select=opportunityid&`$top=1"

            if ($existing.value -and $existing.value.Count -gt 0) {
                $opportunityIds[$opp.key] = $existing.value[0].opportunityid
                Write-Host "    [EXISTS] $($opp.name)" -ForegroundColor DarkGray
                continue
            }

            $closeDate = $weekStart.AddDays($opp.closeDateOffset).ToString("yyyy-MM-dd")

            $body = @{
                name                      = $opp.name
                estimatedvalue            = $opp.amount
                closeprobability          = $opp.probability
                estimatedclosedate        = $closeDate
                description               = $opp.description
                stepname                  = $opp.stage
            }

            # Link to parent account
            $parentAccountId = $accountIds[$opp.account]
            if ($parentAccountId) {
                $body["parentaccountid@odata.bind"] = "/accounts($parentAccountId)"
            }

            # Link primary contact (first contact for the account)
            $primaryContact = $D365Records.contacts | Where-Object { $_.account -eq $opp.account } | Select-Object -First 1
            if ($primaryContact -and $contactIds[$primaryContact.key]) {
                $body["parentcontactid@odata.bind"] = "/contacts($($contactIds[$primaryContact.key]))"
            }

            $result = Invoke-DataverseRequest -Connection $Connection -Method POST -Uri "opportunities" -Body $body
            $opportunityIds[$opp.key] = $result.opportunityid
            Write-Host "    [OK] $($opp.name) - `$$([math]::Round($opp.amount / 1000000, 1))M ($($opp.stage))" -ForegroundColor Green
            $created++
        } catch {
            Write-Host "    [FAIL] $($opp.name): $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }
    }
    Write-Host "  [OPPORTUNITIES] $created created, $failed failed." -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Yellow' })

    # ── Notes / Annotations ──────────────────────────────────────────────────

    $totalNotes = ($D365Records.opportunities | ForEach-Object { $_.notes } | Measure-Object).Count
    if ($totalNotes -gt 0) {
        Write-Host ""
        Write-Host "  NOTES:" -ForegroundColor White
        $created = 0; $failed = 0
        foreach ($opp in $D365Records.opportunities) {
            $oppId = $opportunityIds[$opp.key]
            if (-not $oppId) { continue }

            foreach ($note in $opp.notes) {
                try {
                    $noteDate = $weekStart.AddDays($note.dayOffset).ToString("yyyy-MM-ddT09:00:00Z")
                    $body = @{
                        subject       = $note.subject
                        notetext      = $note.body
                        "objectid_opportunity@odata.bind" = "/opportunities($oppId)"
                    }

                    Invoke-DataverseRequest -Connection $Connection -Method POST -Uri "annotations" -Body $body | Out-Null
                    Write-Host "    [OK] $($note.subject)" -ForegroundColor Green
                    $created++
                } catch {
                    Write-Host "    [FAIL] $($note.subject): $($_.Exception.Message)" -ForegroundColor Red
                    $failed++
                }
            }
        }
        Write-Host "  [NOTES] $created created, $failed failed." -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Yellow' })
    }
}

function Reset-DemoD365 {
    param(
        [Parameter(Mandatory)][hashtable]$Connection,
        [Parameter(Mandatory)]$D365Records,
        [switch]$WhatIf
    )

    $deleted = 0

    # Delete opportunities first (dependent on accounts)
    Write-Host "  OPPORTUNITIES:" -ForegroundColor White
    foreach ($opp in $D365Records.opportunities) {
        try {
            $existing = Invoke-DataverseRequest -Connection $Connection -Method GET `
                -Uri "opportunities?`$filter=name eq '$($opp.name -replace "'","''")'&`$select=opportunityid&`$top=1"

            foreach ($item in $existing.value) {
                # Delete associated notes first
                try {
                    $notes = Invoke-DataverseRequest -Connection $Connection -Method GET `
                        -Uri "annotations?`$filter=_objectid_value eq $($item.opportunityid)&`$select=annotationid"
                    foreach ($n in $notes.value) {
                        if ($WhatIf) {
                            Write-Host "    [WOULD DELETE] Note on $($opp.name)" -ForegroundColor Yellow
                        } else {
                            Invoke-DataverseRequest -Connection $Connection -Method DELETE -Uri "annotations($($n.annotationid))"
                            Write-Host "    [DELETED] Note on $($opp.name)" -ForegroundColor Green
                        }
                    }
                } catch { }

                if ($WhatIf) {
                    Write-Host "    [WOULD DELETE] $($opp.name)" -ForegroundColor Yellow
                } else {
                    Invoke-DataverseRequest -Connection $Connection -Method DELETE -Uri "opportunities($($item.opportunityid))"
                    Write-Host "    [DELETED] $($opp.name)" -ForegroundColor Green
                }
                $deleted++
            }
        } catch {
            Write-Host "    [FAIL] $($opp.name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Delete contacts
    Write-Host "  CONTACTS:" -ForegroundColor White
    foreach ($contact in $D365Records.contacts) {
        try {
            $existing = Invoke-DataverseRequest -Connection $Connection -Method GET `
                -Uri "contacts?`$filter=firstname eq '$($contact.firstName)' and lastname eq '$($contact.lastName)'&`$select=contactid&`$top=1"

            foreach ($item in $existing.value) {
                if ($WhatIf) {
                    Write-Host "    [WOULD DELETE] $($contact.firstName) $($contact.lastName)" -ForegroundColor Yellow
                } else {
                    Invoke-DataverseRequest -Connection $Connection -Method DELETE -Uri "contacts($($item.contactid))"
                    Write-Host "    [DELETED] $($contact.firstName) $($contact.lastName)" -ForegroundColor Green
                }
                $deleted++
            }
        } catch {
            Write-Host "    [FAIL] $($contact.firstName) $($contact.lastName): $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Delete accounts last
    Write-Host "  ACCOUNTS:" -ForegroundColor White
    foreach ($acct in $D365Records.accounts) {
        try {
            $existing = Invoke-DataverseRequest -Connection $Connection -Method GET `
                -Uri "accounts?`$filter=name eq '$($acct.name -replace "'","''")'&`$select=accountid&`$top=1"

            foreach ($item in $existing.value) {
                if ($WhatIf) {
                    Write-Host "    [WOULD DELETE] $($acct.name)" -ForegroundColor Yellow
                } else {
                    Invoke-DataverseRequest -Connection $Connection -Method DELETE -Uri "accounts($($item.accountid))"
                    Write-Host "    [DELETED] $($acct.name)" -ForegroundColor Green
                }
                $deleted++
            }
        } catch {
            Write-Host "    [FAIL] $($acct.name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Write-Host "  [D365] $deleted items $(if ($WhatIf) {'would be '})deleted." -ForegroundColor $(if ($deleted -gt 0) { 'Green' } else { 'DarkGray' })
}

function Get-IndustryCode {
    param([string]$Industry)
    $map = @{
        "Manufacturing" = 12
        "Technology"    = 33
        "Healthcare"    = 14
        "Financial"     = 10
        "Retail"        = 25
        "Education"     = 7
    }
    if ($map.ContainsKey($Industry)) { return $map[$Industry] }
    return $null
}
