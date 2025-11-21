<#
CIS Microsoft 365 Foundations v5.0.0
Control 1.1.3 (L1): Ensure that between two and four global admins are designated

This script:
  - Connects to Microsoft Graph
  - Identifies the "Global Administrator" directory role
  - Enumerates all active user members of that role
  - Outputs the list and checks whether the count is between 2 and 4
  - Emits PASS/FAIL and sets $global:CISCheckResult for pipeline use

REQUIRES:
  - Microsoft Graph PowerShell SDK
      Install-Module Microsoft.Graph -Scope CurrentUser
  - Permissions (consent in your tenant):
      Directory.Read.All
      RoleManagement.Read.Directory
#>

param(
    [string]$ReportPath
)

# -----------------------------------------------------------
# 1. Connect to Microsoft Graph
# -----------------------------------------------------------
$scopes = @(
    "Directory.Read.All",
    "RoleManagement.Read.Directory"
)

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes $scopes


# -----------------------------------------------------------
# 2. Locate the Global Administrator role
# -----------------------------------------------------------
Write-Host "Retrieving directory roles..." -ForegroundColor Cyan

$allRoles = Get-MgDirectoryRole -All

# Primary match by display name
$globalAdminRole = $allRoles | Where-Object { $_.DisplayName -eq "Global Administrator" } | Select-Object -First 1

# Fallback: some tenants may still show "Company Administrator"
if (-not $globalAdminRole) {
    $globalAdminRole = $allRoles | Where-Object { $_.DisplayName -eq "Company Administrator" } | Select-Object -First 1
}

if (-not $globalAdminRole) {
    Write-Error "Could not find the Global Administrator (or Company Administrator) role in this tenant."
    $global:CISCheckResult = "FAIL"
    return
}

Write-Host ("Found Global Admin role: {0} ({1})" -f $globalAdminRole.DisplayName, $globalAdminRole.Id) -ForegroundColor Cyan


# -----------------------------------------------------------
# 3. Enumerate Global Admin members (users only)
# -----------------------------------------------------------
Write-Host "`nEnumerating members of the Global Administrator role..." -ForegroundColor Cyan

$rawMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $globalAdminRole.Id -All -ErrorAction SilentlyContinue

if (-not $rawMembers) {
    Write-Warning "No members found in the Global Administrator role."
    $global:CISCheckResult = "FAIL"
    return
}

$globalAdmins = @()

foreach ($member in $rawMembers) {
    # Try to resolve as a user; skip non-user objects (groups, service principals)
    $user = $null
    try {
        $user = Get-MgUser -UserId $member.Id `
            -Property Id,DisplayName,UserPrincipalName,AccountEnabled `
            -ErrorAction Stop
    }
    catch {
        continue
    }

    if ([string]::IsNullOrWhiteSpace($user.UserPrincipalName)) {
        continue
    }

    # Keep enabled & disabled – CIS doesn’t say to ignore disabled, but we’ll show it
    $globalAdmins += [pscustomobject]@{
        Control     = "1.1.3 (L1)"
        DisplayName = $user.DisplayName
        UPN         = $user.UserPrincipalName
        AccountEnabled = $user.AccountEnabled
    }
}

# Deduplicate by UPN just in case
$globalAdmins = $globalAdmins | Sort-Object UPN -Unique

$gaCount = $globalAdmins.Count


# -----------------------------------------------------------
# 4. Output summary & PASS/FAIL
# -----------------------------------------------------------
Write-Host ""
Write-Host "===== CIS 1.1.3 (L1) – Global Administrator Count Check =====" -ForegroundColor Cyan
Write-Host "Total Global Administrator accounts found : $gaCount"
Write-Host "Recommended range (inclusive)            : 2 to 4" -ForegroundColor Cyan
Write-Host "=============================================================" -ForegroundColor Cyan

if ($gaCount -gt 0) {
    $globalAdmins | Sort-Object DisplayName | Format-Table -AutoSize
} else {
    Write-Host "No Global Administrator accounts resolved as users." -ForegroundColor Yellow
}

# CIS logic: between 2 and 4 inclusive
if ($gaCount -ge 2 -and $gaCount -le 4) {
    Write-Host "`nRESULT: PASS – Number of Global Administrators is within the recommended range (2–4)." -ForegroundColor Green
    $global:CISCheckResult = "PASS"
} else {
    Write-Host "`nRESULT: FAIL – Number of Global Administrators is outside the recommended range (2–4)." -ForegroundColor Red
    $global:CISCheckResult = "FAIL"
}

# -----------------------------------------------------------
# 5. Optional CSV export
# -----------------------------------------------------------
if ($ReportPath) {
    $globalAdmins | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nGlobal Administrator list exported to: $ReportPath" -ForegroundColor Cyan
}

Write-Host "`nCIS Control 1.1.3 check complete.`n" -ForegroundColor Cyan
