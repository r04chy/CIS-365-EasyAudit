<#
CIS Microsoft 365 Foundations v5.0.0
Control 1.3.7 (L2): Ensure 'third-party storage services' are restricted in
'Microsoft 365 on the web' (Automated)

Description:
Third-party storage can be enabled for users in Microsoft 365, allowing them to
store and share documents using services such as Dropbox alongside OneDrive
and SharePoint Online.

CIS requirement:
Restrict third-party storage services in Microsoft 365 on the web.

CIS PowerShell guidance (summarized):
- Check the service principal with appId:
      c1f33bc0-bdb4-4248-ba9b-096807ddb43e
- If the SP is missing OR AccountEnabled = $true -> FAIL
- Only compliant when AccountEnabled = $false

This script:
  - Connects to Microsoft Graph
  - Locates the relevant service principal
  - Evaluates its existence and AccountEnabled state
  - Outputs PASS/FAIL and sets $global:CISCheckResult
  - Optionally exports a CSV report

REQUIRES:
  - Microsoft Graph PowerShell SDK:
        Install-Module Microsoft.Graph -Scope CurrentUser
  - Permissions:
        Application.Read.All   (for audit/read)
#>

[CmdletBinding()]
param(
    [string]$ReportPath
)

# -----------------------------------------------------------
# 1. Connect to Microsoft Graph
# -----------------------------------------------------------
$scopes = @("Application.Read.All")

Write-Host "Connecting to Microsoft Graph with scope: $($scopes -join ', ')" -ForegroundColor Cyan

try {
    if (-not (Get-MgContext)) {
        Connect-MgGraph -Scopes $scopes -ErrorAction Stop
    } else {
        Write-Host "Existing Graph context detected. Ensure it includes Application.Read.All." -ForegroundColor Yellow
    }
}
catch {
    Write-Error "Failed to connect to Microsoft Graph. Error: $_"
    return
}

# -----------------------------------------------------------
# 2. Retrieve the target Service Principal
#    appId = c1f33bc0-bdb4-4248-ba9b-096807ddb43e
# -----------------------------------------------------------
$targetAppId = "c1f33bc0-bdb4-4248-ba9b-096807ddb43e"
Write-Host "Retrieving service principal for appId $targetAppId..." -ForegroundColor Cyan

$sp = $null
try {
    $sp = Get-MgServicePrincipal -Filter "appId eq '$targetAppId'" -ErrorAction Stop
}
catch {
    Write-Error "Failed to query service principals from Graph. Error: $_"
    $global:CISCheckResult = "FAIL"
    return
}

$exists          = $false
$accountEnabled  = $null
$compliant       = $false
$notes           = $null
$spId            = $null
$spDisplayName   = $null

if (-not $sp) {
    # Per CIS: if SP doesn't exist, users can still use third-party storage -> FAIL
    $exists = $false
    $accountEnabled = $null
    $compliant = $false
    $notes = "Service principal missing; third-party storage services remain available."
} else {
    $exists         = $true
    $spId           = $sp.Id
    $spDisplayName  = $sp.DisplayName
    $accountEnabled = $sp.AccountEnabled

    if ($accountEnabled -eq $false) {
        $compliant = $true
        $notes = "Service principal exists and is disabled; third-party storage services are restricted."
    } else {
        $compliant = $false
        $notes = "Service principal is enabled; users can open files stored in third-party storage services."
    }
}

# -----------------------------------------------------------
# 3. Build result object
# -----------------------------------------------------------
$result = [pscustomobject]@{
    Control              = "1.3.7 (L2)"
    Setting              = "Third-party storage services in Microsoft 365 on the web"
    ServicePrincipalAppId= $targetAppId
    ServicePrincipalId   = $spId
    ServicePrincipalName = $spDisplayName
    Exists               = $exists
    AccountEnabled       = $accountEnabled
    Compliant            = $compliant
    Notes                = $notes
}

# -----------------------------------------------------------
# 4. Output summary & PASS/FAIL
# -----------------------------------------------------------
Write-Host ""
Write-Host "===== CIS 1.3.7 (L2) – Third-party Storage Services in Microsoft 365 on the web =====" -ForegroundColor Cyan
$result | Format-Table Control, ServicePrincipalName, Exists, AccountEnabled, Compliant -AutoSize
Write-Host "CIS requirement: Service principal must exist AND AccountEnabled = False" -ForegroundColor Cyan
Write-Host "===============================================================================" -ForegroundColor Cyan

if ($compliant) {
    Write-Host "`nRESULT: PASS – Third-party storage services are restricted in Microsoft 365 on the web." -ForegroundColor Green
    $global:CISCheckResult = "PASS"
}
else {
    Write-Host "`nRESULT: FAIL – Third-party storage services are NOT fully restricted in Microsoft 365 on the web." -ForegroundColor Red
    $global:CISCheckResult = "FAIL"
}

# -----------------------------------------------------------
# 5. Optional CSV export
# -----------------------------------------------------------
if ($ReportPath) {
    try {
        $result | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
        Write-Host "`nThird-party storage services report exported to: $ReportPath" -ForegroundColor Cyan
    }
    catch {
        Write-Warning "Failed to export CSV report. Error: $_"
    }
}

Write-Host "`nCIS Control 1.3.7 check complete.`n" -ForegroundColor Cyan

<#
REFERENCE REMEDIATION (NOT run by this script):

To remediate via PowerShell per CIS:

    Connect-MgGraph -Scopes "Application.ReadWrite.All"

    $targetAppId = "c1f33bc0-bdb4-4248-ba9b-096807ddb43e"
    $sp = Get-MgServicePrincipal -Filter "appId eq '$targetAppId'"

    if (-not $sp) {
        $sp = New-MgServicePrincipal -AppId $targetAppId
    }

    Update-MgServicePrincipal -ServicePrincipalId $sp.Id -AccountEnabled:$false

#>
