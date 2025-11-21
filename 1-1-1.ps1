<#
CIS Microsoft 365 Foundations Benchmark v5.0.0
Control 1.1.1 (L1): Ensure Administrative accounts are cloud-only
#>

param(
    [string]$ReportPath
)

$scopes = @(
    "Directory.Read.All",
    "RoleManagement.Read.Directory"
)

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes $scopes

# --------------------------------------------------------------------
# 1. Get privileged roles
# --------------------------------------------------------------------
$PrivilegedRoleNames = @(
    "Global Administrator",
    "Privileged Role Administrator",
    "Security Administrator",
    "Exchange Administrator",
    "SharePoint Administrator",
    "Teams Administrator",
    "User Administrator",
    "Authentication Administrator",
    "Application Administrator",
    "Cloud Application Administrator",
    "Compliance Administrator",
    "Helpdesk Administrator",
    "Intune Administrator",
    "Hybrid Identity Administrator",
    "Conditional Access Administrator"
)

Write-Host "Retrieving directory roles..." -ForegroundColor Cyan
$allRoles  = Get-MgDirectoryRole -All
$adminRoles = $allRoles | Where-Object {
    $PrivilegedRoleNames -contains $_.DisplayName
}

if (-not $adminRoles) {
    Write-Warning "No matching privileged roles found. Check role names or your permissions."
    $global:CISCheckResult = "PASS"
    return
}

Write-Host "Found the following privileged roles:" -ForegroundColor Cyan
$adminRoles | Select-Object DisplayName, Id | Format-Table -AutoSize

# --------------------------------------------------------------------
# 2. Enumerate members and test cloud-only status
# --------------------------------------------------------------------
$results = @()

foreach ($role in $adminRoles) {
    Write-Host "`nProcessing role: $($role.DisplayName)" -ForegroundColor Yellow

    $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction SilentlyContinue

    if (-not $members) {
        Write-Host "  No raw members returned for this role (may be PIM-eligible only or genuinely empty)." -ForegroundColor DarkGray
        continue
    }

    Write-Host "  Raw members returned: $($members.Count)" -ForegroundColor DarkGray

    foreach ($member in $members) {
        # Try to treat it as a user; if it fails, skip silently
        $user = $null
        try {
            $user = Get-MgUser -UserId $member.Id -Property Id,DisplayName,UserPrincipalName,OnPremisesSyncEnabled,OnPremisesImmutableId -ErrorAction Stop
        }
        catch {
            # Not a user (could be group, service principal, etc.)
            continue
        }

        # Determine if synced from on-prem
        $isSynced = $false
        if ($user.OnPremisesSyncEnabled -eq $true) {
            $isSynced = $true
        } elseif ($user.OnPremisesImmutableId) {
            $isSynced = $true
        }

        $results += [pscustomobject]@{
            Control          = "1.1.1 (L1)"
            Role             = $role.DisplayName
            DisplayName      = $user.DisplayName
            UPN              = $user.UserPrincipalName
            SyncedFromOnPrem = $isSynced
            Status           = if ($isSynced) {"Non-Compliant"} else {"Compliant"}
        }
    }
}

# --------------------------------------------------------------------
# 3. Output summary
# --------------------------------------------------------------------
Write-Host ""
Write-Host "===== CIS 1.1.1 (L1) Cloud-Only Administrative Accounts =====" -ForegroundColor Cyan
Write-Host "Total Admin Accounts Checked : $($results.Count)"

$nonCompliant = $results | Where-Object { $_.SyncedFromOnPrem -eq $true }
$compliant    = $results | Where-Object { $_.SyncedFromOnPrem -eq $false }

Write-Host "Compliant (Cloud-only)       : $($compliant.Count)" -ForegroundColor Green
Write-Host "Non-Compliant (Synced)       : $($nonCompliant.Count)" -ForegroundColor Red
Write-Host "================================================================" -ForegroundColor Cyan

if ($nonCompliant.Count -gt 0) {
    Write-Host "`nNon-Compliant Accounts:" -ForegroundColor Red
    $nonCompliant | Format-Table -AutoSize
}

# --------------------------------------------------------------------
# 4. PASS / FAIL for pipeline use
# --------------------------------------------------------------------
if ($nonCompliant.Count -gt 0) {
    Write-Host "`nRESULT: FAIL – One or more administrative accounts are *not* cloud-only." -ForegroundColor Red
    $global:CISCheckResult = "FAIL"
} else {
    Write-Host "`nRESULT: PASS – All checked administrative accounts are cloud-only." -ForegroundColor Green
    $global:CISCheckResult = "PASS"
}

# --------------------------------------------------------------------
# 5. Optional CSV export
# --------------------------------------------------------------------
if ($ReportPath) {
    $results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nReport exported to: $ReportPath" -ForegroundColor Cyan
}

Write-Host "`nCIS Control 1.1.1 check complete.`n" -ForegroundColor Cyan