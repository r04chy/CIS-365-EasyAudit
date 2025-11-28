<#
CIS 1.3.5 – Ensure internal phishing protection for Forms is enabled
Updated script that handles tenants where the Forms admin API is not available.
#>

[CmdletBinding()]
param(
    [string]$ReportPath
)

$scopes = @("OrgSettings.Read.All")

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    if (-not (Get-MgContext)) {
        Connect-MgGraph -Scopes $scopes -ErrorAction Stop
    }
}
catch {
    Write-Error "Failed to connect to Graph: $_"
    return
}

$uri = "https://graph.microsoft.com/beta/admin/forms/settings"

Write-Host "Checking Forms phishing protection..." -ForegroundColor Cyan

try {
    $settings = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
}
catch {
    Write-Warning "Forms admin API is NOT available in this tenant."
    Write-Warning "This tenant likely does NOT have Forms phishing protection rolled out yet."
    
    $result = [pscustomobject]@{
        Control     = "1.3.5 (L1)"
        Supported   = $false
        Compliant   = $false
        Reason      = "Forms admin API not available in this tenant"
    }

    if ($ReportPath) {
        $result | Export-Csv -Path $ReportPath -NoTypeInformation
    }

    $global:CISCheckResult = "FAIL"
    return
}

$phishingEnabled = $settings.isInOrgFormsPhishingScanEnabled

$result = [pscustomobject]@{
    Control     = "1.3.5 (L1)"
    Supported   = $true
    PhishingEnabled = $phishingEnabled
    Compliant   = ($phishingEnabled -eq $true)
}

$result | Format-Table -AutoSize

if ($phishingEnabled -eq $true) {
    Write-Host "`nPASS – Internal phishing protection is enabled." -ForegroundColor Green
    $global:CISCheckResult = "PASS"
} else {
    Write-Host "`nFAIL – Internal phishing protection is disabled." -ForegroundColor Red
    $global:CISCheckResult = "FAIL"
}

if ($ReportPath) {
    $result | Export-Csv -Path $ReportPath -NoTypeInformation
}
