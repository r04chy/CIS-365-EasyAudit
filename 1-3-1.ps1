<#
CIS Microsoft 365 Foundations v5.0.0
Control 1.3.1 (L1): Ensure the 'Password expiration policy' is set to
                    'Set passwords to never expire (recommended)'

This script:
  - Uses Microsoft Graph to read domain password policies
  - Checks PasswordValidityPeriodInDays for each verified domain
  - Considers a domain COMPLIANT if PasswordValidityPeriodInDays = 2147483647
    (equivalent to "Set passwords to never expire (recommended)")
  - Outputs:
      * Per-domain password validity settings
      * PASS/FAIL for CIS 1.3.1
  - Optional CSV export

REQUIRES:
  - Microsoft Graph PowerShell SDK:
        Install-Module Microsoft.Graph -Scope CurrentUser
  - Permissions:
        Domain.Read.All
#>

param(
    [string]$ReportPath
)

# -----------------------------------------------------------
# 1. Connect to Microsoft Graph
# -----------------------------------------------------------
$scopes = @("Domain.Read.All")

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes $scopes -ErrorAction Stop
}
catch {
    Write-Error "Failed to connect to Microsoft Graph. Error: $_"
    return
}

# -----------------------------------------------------------
# 2. Retrieve domains and their password policies
# -----------------------------------------------------------
Write-Host "Retrieving Microsoft 365 domains and password policies..." -ForegroundColor Cyan

try {
    $domains = Get-MgDomain -All -ErrorAction Stop
}
catch {
    Write-Error "Failed to retrieve domains from Microsoft Graph. Error: $_"
    return
}

if (-not $domains) {
    Write-Host "No domains returned from Microsoft Graph." -ForegroundColor Yellow
    $global:CISCheckResult = "PASS"
    return
}

# Only care about verified domains (typical CIS intent)
$targetDomains = $domains | Where-Object { $_.IsVerified -eq $true }

if (-not $targetDomains) {
    Write-Host "No verified domains found in this tenant." -ForegroundColor Yellow
    $global:CISCheckResult = "PASS"
    return
}

$results = @()

foreach ($d in $targetDomains) {
    # PasswordValidityPeriodInDays:
    #   2147483647 = 'Set passwords to never expire (recommended)'
    #   Null/0/other = not compliant with CIS 1.3.1
    $validity = $d.PasswordValidityPeriodInDays
    $notify   = $d.PasswordNotificationWindowInDays

    $isCompliant = $false
    if ($validity -eq 2147483647) {
        $isCompliant = $true
    }

    $results += [pscustomobject]@{
        Control                     = "1.3.1 (L1)"
        DomainName                  = $d.Id
        IsVerified                  = $d.IsVerified
        PasswordValidityPeriodInDays= $validity
        PasswordNotificationWindowInDays = $notify
        Compliant                   = $isCompliant
    }
}

# -----------------------------------------------------------
# 3. Output summary & CIS evaluation
# -----------------------------------------------------------
$nonCompliant = $results | Where-Object { $_.Compliant -eq $false }
$compliant    = $results | Where-Object { $_.Compliant -eq $true }

Write-Host ""
Write-Host "===== CIS 1.3.1 (L1) – Password Expiration Policy =====" -ForegroundColor Cyan
Write-Host "Verified domains checked                    : $($results.Count)"
Write-Host "Domains with 'never expire' configured      : $($compliant.Count)" -ForegroundColor Green
Write-Host "Domains without 'never expire' configured   : $($nonCompliant.Count)" -ForegroundColor Red
Write-Host "========================================================" -ForegroundColor Cyan

$results | Sort-Object DomainName | Format-Table `
    DomainName, IsVerified, PasswordValidityPeriodInDays, PasswordNotificationWindowInDays, Compliant -AutoSize

if ($nonCompliant.Count -gt 0) {
    Write-Host "`nNon-compliant domains (passwords ARE set to expire):" -ForegroundColor Red
    $nonCompliant | Sort-Object DomainName | Format-Table DomainName, PasswordValidityPeriodInDays -AutoSize

    Write-Host "`nRESULT: FAIL – One or more verified domains do NOT have 'passwords never expire' configured." -ForegroundColor Red
    $global:CISCheckResult = "FAIL"
}
else {
    Write-Host "`nRESULT: PASS – All verified domains are configured for 'passwords never expire'." -ForegroundColor Green
    $global:CISCheckResult = "PASS"
}

# -----------------------------------------------------------
# 4. Optional CSV export
# -----------------------------------------------------------
if ($ReportPath) {
    try {
        $results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
        Write-Host "`nPassword expiration policy report exported to: $ReportPath" -ForegroundColor Cyan
    }
    catch {
        Write-Warning "Failed to export CSV report. Error: $_"
    }
}

Write-Host "`nCIS Control 1.3.1 check complete.`n" -ForegroundColor Cyan
