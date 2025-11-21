<#
CIS Microsoft 365 Foundations v5.0.0
Control 1.2.2 (L1): Ensure sign-in to shared mailboxes is blocked (Automated)

This script:
  - Connects to Exchange Online
  - Retrieves all shared mailboxes
  - Checks whether sign-in is blocked for each mailbox by inspecting AccountDisabled
  - Outputs:
      * Shared mailbox list with AccountDisabled status
      * PASS/FAIL for CIS 1.2.2
  - Optional CSV export

REQUIRES:
  - Exchange Online PowerShell V3 module:
        Install-Module ExchangeOnlineManagement -Scope CurrentUser
  - Permissions:
        Exchange Online admin role (e.g. Exchange Administrator)
#>

param(
    [string]$ReportPath
)

# -----------------------------------------------------------
# 1. Connect to Exchange Online
# -----------------------------------------------------------
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan

try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
}
catch {
    Write-Error "ExchangeOnlineManagement module not found. Install it with: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    return
}

try {
    Connect-ExchangeOnline -ShowProgress $false
}
catch {
    Write-Error "Failed to connect to Exchange Online. Check your permissions and network connectivity."
    return
}

# -----------------------------------------------------------
# 2. Retrieve shared mailboxes and associated user objects
# -----------------------------------------------------------
Write-Host "Retrieving shared mailboxes..." -ForegroundColor Cyan

try {
    $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
}
catch {
    Write-Error "Error retrieving shared mailboxes: $_"
    return
}

if (-not $sharedMailboxes) {
    Write-Host "No shared mailboxes found in this tenant." -ForegroundColor Yellow
    $global:CISCheckResult = "PASS"
    Disconnect-ExchangeOnline -Confirm:$false
    return
}

Write-Host ("Total shared mailboxes found: {0}" -f $sharedMailboxes.Count) -ForegroundColor Cyan

$results = @()

foreach ($mbx in $sharedMailboxes) {
    # Get the corresponding user object to check AccountDisabled / sign-in status
    $user = $null
    try {
        $user = Get-User -Identity $mbx.Identity -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not retrieve user object for mailbox: $($mbx.PrimarySmtpAddress)"
        continue
    }

    $results += [pscustomobject]@{
        Control          = "1.2.2 (L1)"
        DisplayName      = $mbx.DisplayName
        PrimarySmtpAddress = $mbx.PrimarySmtpAddress
        RecipientTypeDetails = $mbx.RecipientTypeDetails
        AccountDisabled  = $user.AccountDisabled
    }
}

# -----------------------------------------------------------
# 3. Evaluate CIS 1.2.2 & output
# -----------------------------------------------------------
$nonCompliant = $results | Where-Object { $_.AccountDisabled -ne $true }
$compliant    = $results | Where-Object { $_.AccountDisabled -eq $true }

Write-Host ""
Write-Host "===== CIS 1.2.2 (L1) – Shared Mailbox Sign-in Block Check =====" -ForegroundColor Cyan
Write-Host "Total shared mailboxes checked       : $($results.Count)"
Write-Host "Compliant (sign-in blocked)          : $($compliant.Count)" -ForegroundColor Green
Write-Host "Non-compliant (sign-in NOT blocked)  : $($nonCompliant.Count)" -ForegroundColor Red
Write-Host "================================================================" -ForegroundColor Cyan

$results | Sort-Object DisplayName | Format-Table DisplayName,PrimarySmtpAddress,RecipientTypeDetails,AccountDisabled -AutoSize

if ($nonCompliant.Count -gt 0) {
    Write-Host "`nNon-compliant shared mailboxes (AccountDisabled = False or null):" -ForegroundColor Red
    $nonCompliant | Sort-Object DisplayName | Format-Table DisplayName,PrimarySmtpAddress,AccountDisabled -AutoSize

    Write-Host "`nRESULT: FAIL – One or more shared mailboxes allow sign-in." -ForegroundColor Red
    $global:CISCheckResult = "FAIL"
} else {
    Write-Host "`nRESULT: PASS – All shared mailboxes have sign-in blocked (AccountDisabled = True)." -ForegroundColor Green
    $global:CISCheckResult = "PASS"
}

# -----------------------------------------------------------
# 4. Optional CSV export
# -----------------------------------------------------------
if ($ReportPath) {
    try {
        $results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
        Write-Host "`nShared mailbox sign-in report exported to: $ReportPath" -ForegroundColor Cyan
    }
    catch {
        Write-Warning "Failed to export CSV report: $_"
    }
}

# -----------------------------------------------------------
# 5. Cleanup
# -----------------------------------------------------------
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "`nCIS Control 1.2.2 check complete.`n" -ForegroundColor Cyan
