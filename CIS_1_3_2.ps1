<#
CIS Microsoft 365 Foundations v5.0.0
Control 1.3.2 (L2): Ensure 'Idle session timeout' is set to
                    '3 hours (or less)' for unmanaged devices (Automated)

This script:
  - Auto-detects the SharePoint Online Admin URL using Microsoft Graph
  - Connects to SharePoint Online Admin
  - Uses Get-SPOBrowserIdleSignOut to read the current idle session sign-out settings
  - Considers the control COMPLIANT if:
        Enabled = True
        AND SignOutAfter <= 3 hours (180 minutes)
  - Outputs:
        * Enabled / WarnAfter / SignOutAfter values
        * PASS/FAIL for CIS 1.3.2
  - Optionally exports a CSV with the configuration

REQUIRES:
  - Microsoft Graph PowerShell SDK:
        Install-Module Microsoft.Graph -Scope CurrentUser
  - SharePoint Online Management Shell:
        Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
  - Permissions:
        Directory.Read.All (for auto-detection via Graph)
        SharePoint Administrator (for SPO connection)
#>

[CmdletBinding()]
param(
    # Optional: override auto-detected admin center URL, e.g. https://contoso-admin.sharepoint.com
    [string]$AdminCenterUrl,

    [string]$ReportPath
)

# -----------------------------------------------------------
# 0. Helper: Auto-detect SPO Admin URL using Microsoft Graph
# -----------------------------------------------------------
function Get-AutoDetectedSPOAdminUrl {
    Write-Host "Auto-detecting SharePoint Online Admin URL via Microsoft Graph..." -ForegroundColor Cyan

    $graphScopes = @("Directory.Read.All")

    try {
        # Reuse existing Graph session if already connected, otherwise connect
        if (-not (Get-MgContext)) {
            Connect-MgGraph -Scopes $graphScopes -ErrorAction Stop
        }
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph for auto-detection. Error: $_"
        return $null
    }

    try {
        $org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
    }
    catch {
        Write-Error "Failed to retrieve organization information from Microsoft Graph. Error: $_"
        return $null
    }

    if (-not $org.VerifiedDomains) {
        Write-Error "No VerifiedDomains found on the organization object."
        return $null
    }

    # Prefer the initial domain, e.g. contoso.onmicrosoft.com
    $initialDomain = $org.VerifiedDomains | Where-Object { $_.IsInitial -eq $true } | Select-Object -First 1
    if (-not $initialDomain) {
        # Fallback: just pick the first verified domain
        $initialDomain = $org.VerifiedDomains | Select-Object -First 1
    }

    $domainName = $initialDomain.Name
    if ([string]::IsNullOrWhiteSpace($domainName)) {
        Write-Error "Initial domain name is empty; cannot derive SPO admin URL."
        return $null
    }

    # tenantname = leftmost label from contoso.onmicrosoft.com
    $tenantName = $domainName.Split(".")[0]

    $autoUrl = "https://{0}-admin.sharepoint.com" -f $tenantName
    Write-Host "Detected tenant initial domain: $domainName" -ForegroundColor DarkCyan
    Write-Host "Derived SharePoint Admin URL:  $autoUrl" -ForegroundColor DarkCyan

    return $autoUrl
}

# -----------------------------------------------------------
# 1. Determine AdminCenterUrl (auto-detect if not supplied)
# -----------------------------------------------------------
if (-not $AdminCenterUrl) {
    $AdminCenterUrl = Get-AutoDetectedSPOAdminUrl
    if (-not $AdminCenterUrl) {
        Write-Error "Unable to determine SharePoint Online Admin URL automatically. Supply -AdminCenterUrl explicitly."
        return
    }
}
else {
    Write-Host "Using provided SharePoint Online Admin URL: $AdminCenterUrl" -ForegroundColor Cyan
}

# -----------------------------------------------------------
# 2. Load SPO module & connect
# -----------------------------------------------------------
Write-Host "Loading SharePoint Online PowerShell module..." -ForegroundColor Cyan
try {
    Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop
}
catch {
    Write-Error "Microsoft.Online.SharePoint.PowerShell module not found. Install it with:`n  Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser"
    return
}

Write-Host "Connecting to SharePoint Online Admin: $AdminCenterUrl" -ForegroundColor Cyan
try {
    Connect-SPOService -Url $AdminCenterUrl -ErrorAction Stop
}
catch {
    Write-Error "Failed to connect to SharePoint Online Admin. Check the URL and your permissions. Error: $_"
    return
}

# -----------------------------------------------------------
# 3. Retrieve Idle Session Sign-out configuration
# -----------------------------------------------------------
Write-Host "Retrieving Idle session sign-out configuration..." -ForegroundColor Cyan

try {
    $config = Get-SPOBrowserIdleSignOut
}
catch {
    Write-Error "Failed to retrieve Idle session sign-out settings. Error: $_"
    Disconnect-SPOService
    return
}

if (-not $config) {
    Write-Host "No Idle session sign-out configuration returned. Assuming default (disabled)." -ForegroundColor Yellow
}

$enabled        = $config.Enabled
$warnMinutes    = $null
$signOutMinutes = $null

if ($config.WarnAfter) {
    $warnMinutes = [int]$config.WarnAfter.TotalMinutes
}
if ($config.SignOutAfter) {
    $signOutMinutes = [int]$config.SignOutAfter.TotalMinutes
}

# -----------------------------------------------------------
# 4. Evaluate CIS 1.3.2
#    Requirement: Enabled = True AND SignOutAfter <= 180 minutes (3 hours)
# -----------------------------------------------------------
$compliant = $false

if ($enabled -eq $true -and $signOutMinutes -ne $null -and $signOutMinutes -le 180) {
    $compliant = $true
}

$result = [pscustomobject]@{
    Control                 = "1.3.2 (L2)"
    AdminCenterUrl          = $AdminCenterUrl
    Enabled                 = $enabled
    WarnAfterMinutes        = $warnMinutes
    SignOutAfterMinutes     = $signOutMinutes
    Compliant               = $compliant
}

# -----------------------------------------------------------
# 5. Output summary & PASS/FAIL
# -----------------------------------------------------------
Write-Host ""
Write-Host "===== CIS 1.3.2 (L2) – Idle Session Timeout (Unmanaged Devices) =====" -ForegroundColor Cyan
Write-Host ("Admin Center URL           : {0}" -f $AdminCenterUrl)
Write-Host ("Idle session sign-out      : {0}" -f ($enabled -as [string]))
Write-Host ("Warn after (minutes)       : {0}" -f ($warnMinutes -as [string]))
Write-Host ("Sign out after (minutes)   : {0}" -f ($signOutMinutes -as [string]))
Write-Host "Required: Enabled = True AND SignOutAfter <= 180 minutes (3 hours)" -ForegroundColor Cyan
Write-Host "=====================================================================" -ForegroundColor Cyan

$result | Format-Table -AutoSize

if ($compliant) {
    Write-Host "`nRESULT: PASS – Idle session timeout is enabled and configured to 3 hours or less." -ForegroundColor Green
    $global:CISCheckResult = "PASS"
}
else {
    Write-Host "`nRESULT: FAIL – Idle session timeout is either disabled or greater than 3 hours." -ForegroundColor Red
    $global:CISCheckResult = "FAIL"
}

# -----------------------------------------------------------
# 6. Optional CSV export
# -----------------------------------------------------------
if ($ReportPath) {
    try {
        $result | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
        Write-Host "`nIdle session timeout configuration exported to: $ReportPath" -ForegroundColor Cyan
    }
    catch {
        Write-Warning "Failed to export CSV report. Error: $_"
    }
}

# -----------------------------------------------------------
# 7. Cleanup
# -----------------------------------------------------------
Disconnect-SPOService
Write-Host "`nCIS Control 1.3.2 check complete.`n" -ForegroundColor Cyan
