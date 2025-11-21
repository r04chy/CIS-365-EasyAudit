<#
CIS Microsoft 365 Foundations v5.0.0
Control 1.3.8 (L2): Ensure that Sways cannot be shared with people outside of your organization

IMPORTANT – MANUAL CONTROL
-----------------------------------------------------------
This CIS control CANNOT be automated.

Microsoft DOES NOT expose the Sway external-sharing setting via:

    - Microsoft Graph API (v1.0 or beta)
    - Exchange Online PowerShell
    - SharePoint Online PowerShell
    - Microsoft 365 Admin PowerShell
    - Any other public API or service principal

The setting ONLY exists inside the Microsoft 365 Admin Center UI:

    Microsoft 365 Admin Center → Settings → Org settings → Sway → Sharing

Because no API exists, CIS marks this control as **MANUAL**.

This helper script:

  • Confirms whether Sway is licensed/enabled in the tenant  
  • Outputs clear instructions for manual auditing  
  • Always sets:
        $global:CISCheckResult = "MANUAL"
  • Allows your automated audit pipeline to handle/control/skip this check cleanly

This script DOES NOT determine PASS/FAIL — a human must check the admin UI.

-----------------------------------------------------------
Manual COMPLIANCE CHECK
-----------------------------------------------------------
Navigate to:

    https://admin.microsoft.com  
      → Settings  
      → Org settings  
      → Sway  
      → Sharing section

Ensure this option is **NOT checked**:

    "Let people in your organization share their sways with people outside your organization"

If NOT checked → COMPLIANT  
If checked     → NOT COMPLIANT
-----------------------------------------------------------
#>

[CmdletBinding()]
param(
    [switch]$SkipGraphCheck,
    [string]$ReportPath
)

$ErrorActionPreference = "Stop"

$result = [pscustomobject]@{
    Control          = "1.3.8 (L2)"
    Setting          = "Sway external sharing"
    Automatable      = $false
    SwayAvailable    = $null
    Compliant        = $null
    CheckMethod      = "Manual (Admin Center UI)"
    Notes            = "No supported API exists to read Sway external sharing settings."
}

# Optional tenant capability detection
if (-not $SkipGraphCheck) {
    Write-Host "Checking for Sway availability via Microsoft Graph..." -ForegroundColor Cyan
    $scopes = @("Directory.Read.All")

    try {
        if (-not (Get-MgContext)) {
            Connect-MgGraph -Scopes $scopes -ErrorAction Stop
        }

        $skus = Get-MgSubscribedSku -All
        $swayPlans = $skus.ServicePlans | Where-Object { $_.ServicePlanName -match 'sway' }

        if ($swayPlans) {
            $result.SwayAvailable = $true
        }
        else {
            $result.SwayAvailable = $false
            $result.Notes = "Sway not detected in license SKUs; if Sway exists for this tenant, check manually in admin portal."
        }
    }
    catch {
        Write-Warning "Unable to retrieve SKU data from Graph. Error: $_"
        $result.SwayAvailable = $null
        $result.Notes = "Unable to confirm Sway availability; perform manual UI check regardless."
    }
}

Write-Host ""
Write-Host "===== CIS 1.3.8 (L2) – Sway External Sharing =====" -ForegroundColor Cyan
Write-Host "This control is MANUAL and cannot be automated." -ForegroundColor Yellow
Write-Host ""
Write-Host "Manual steps:" -ForegroundColor Cyan
Write-Host " 1. Go to https://admin.microsoft.com" -ForegroundColor White
Write-Host " 2. Navigate: Settings → Org settings → Sway" -ForegroundColor White
Write-Host " 3. Under 'Sharing', ensure this option is NOT checked:" -ForegroundColor White
Write-Host "        'Let people in your organization share their sways with people outside your organization'" -ForegroundColor White
Write-Host ""
Write-Host "A human must perform this check manually." -ForegroundColor Cyan
Write-Host ""

$result | Format-Table -AutoSize

$global:CISCheckResult = "MANUAL"

if ($ReportPath) {
    $result | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Results exported to: $ReportPath" -ForegroundColor Cyan
}

Write-Host "`nCIS 1.3.8 check script complete (manual verification required).`n" -ForegroundColor Cyan
