<#
CIS Microsoft 365 Foundations v5.0.0
Control 1.2.1 (L2): Ensure that only organizationally managed/approved public groups exist

This script:
  - Identifies all Microsoft 365 (Unified) Groups
  - Filters groups where "Visibility" = "Public"
  - Outputs public groups for human review
  - Emits PASS/FAIL depending on whether any public groups exist
  - Optional CSV export

REQUIRES:
  - Microsoft Graph PowerShell SDK
  - Permissions:
        Group.Read.All
        Directory.Read.All
#>

param(
    [string]$ReportPath
)

# -----------------------------------------------------------
# 1. Connect to Microsoft Graph
# -----------------------------------------------------------
$scopes = @(
    "Group.Read.All",
    "Directory.Read.All"
)

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes $scopes


# -----------------------------------------------------------
# 2. Retrieve all Microsoft 365 Groups
# -----------------------------------------------------------
Write-Host "Retrieving Microsoft 365 Groups..." -ForegroundColor Cyan

$allGroups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'Unified')" -All -ErrorAction SilentlyContinue

if (-not $allGroups) {
    Write-Host "No Microsoft 365 Groups found in the tenant." -ForegroundColor Yellow
    $global:CISCheckResult = "PASS"
    return
}

Write-Host ("Total Unified (M365) Groups found: {0}" -f $allGroups.Count) -ForegroundColor Cyan


# -----------------------------------------------------------
# 3. Select those with Visibility = "Public"
# -----------------------------------------------------------
$publicGroups = $allGroups | Where-Object { $_.Visibility -eq "Public" }

Write-Host ("Public M365 Groups found: {0}" -f $publicGroups.Count) -ForegroundColor Cyan

$results = @()

foreach ($g in $publicGroups) {
    $owners = @()

    try {
        $ownerObjs = Get-MgGroupOwner -GroupId $g.Id -All -ErrorAction SilentlyContinue
        foreach ($o in $ownerObjs) {
            if ($o.AdditionalProperties["userPrincipalName"]) {
                $owners += $o.AdditionalProperties["userPrincipalName"]
            }
        }
    }
    catch {}

    $results += [pscustomobject]@{
        Control     = "1.2.1 (L2)"
        GroupName   = $g.DisplayName
        GroupId     = $g.Id
        Visibility  = $g.Visibility
        Description = $g.Description
        Owners      = ($owners -join ", ")
    }
}


# -----------------------------------------------------------
# 4. Output & CIS Evaluation
# -----------------------------------------------------------
Write-Host ""
Write-Host "===== CIS 1.2.1 (L2) – Public Microsoft 365 Groups Audit =====" -ForegroundColor Cyan
Write-Host "Total Public Groups: $($publicGroups.Count)" -ForegroundColor Cyan
Write-Host "A HUMAN reviewer must confirm whether each public group is approved." -ForegroundColor Yellow
Write-Host "================================================================" -ForegroundColor Cyan

if ($results.Count -gt 0) {
    $results | Sort-Object GroupName | Format-Table -AutoSize
} else {
    Write-Host "No public Microsoft 365 Groups found." -ForegroundColor Green
}

if ($results.Count -eq 0) {
    Write-Host "`nRESULT: PASS – No public M365 groups detected." -ForegroundColor Green
    $global:CISCheckResult = "PASS"
}
else {
    Write-Host "`nRESULT: FAIL – One or more public M365 groups exist and must be manually approved." -ForegroundColor Red
    $global:CISCheckResult = "FAIL"
}


# -----------------------------------------------------------
# 5. Optional CSV export
# -----------------------------------------------------------
if ($ReportPath) {
    $results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nPublic group report exported to: $ReportPath" -ForegroundColor Cyan
}

Write-Host "`nCIS Control 1.2.1 check complete.`n" -ForegroundColor Cyan
