<#
CIS Microsoft 365 Foundations v5.0.0
Control 1.1.2 (L1): Ensure two emergency access accounts have been defined

This script:
  - Identifies privileged administrative roles
  - Enumerates all admin users
  - Aggregates roles per user
  - Builds a full MFA profile per user using:
      * Authentication methods endpoints (FIDO2, Auth App, Phone, TAP, etc.)
      * Conditional Access policies that require MFA
  - Outputs a table and optional CSV for HUMAN review

REQUIRES:
  - Microsoft Graph PowerShell SDK
      Install-Module Microsoft.Graph -Scope CurrentUser
  - Permissions (tenant admin consent):
      Directory.Read.All
      RoleManagement.Read.Directory
      UserAuthenticationMethod.Read.All
      Policy.Read.ConditionalAccess
#>

param(
    [string]$ReportPath
)

# ------------------------------
# 1. Connect to Microsoft Graph
# ------------------------------
$scopes = @(
    "Directory.Read.All",
    "RoleManagement.Read.Directory",
    "UserAuthenticationMethod.Read.All",
    "Policy.Read.ConditionalAccess"
)

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes $scopes


# ------------------------------
# 2. Define privileged / admin roles
# ------------------------------
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

Write-Host "Retrieving privileged roles..." -ForegroundColor Cyan

$allRoles   = Get-MgDirectoryRole -All
$adminRoles = $allRoles | Where-Object {
    $PrivilegedRoleNames -contains $_.DisplayName
}

if (-not $adminRoles) {
    Write-Warning "No privileged roles found. Exiting."
    return
}

Write-Host "`nPrivileged roles detected:" -ForegroundColor Cyan
$adminRoles | Select-Object DisplayName, Id | Format-Table -AutoSize


# ------------------------------
# 3. Collect raw (user, role) pairs
# ------------------------------
Write-Host "`nEnumerating admin accounts..." -ForegroundColor Cyan

$rawUserRoles = @()

foreach ($role in $adminRoles) {
    Write-Host "`nProcessing role: $($role.DisplayName)" -ForegroundColor Yellow

    $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction SilentlyContinue
    if (-not $members) {
        Write-Host "  No members assigned (PIM-eligible-only or empty)" -ForegroundColor DarkGray
        continue
    }

    foreach ($member in $members) {
        # Try to resolve as a user; skip non-user objects
        $user = $null
        try {
            $user = Get-MgUser -UserId $member.Id `
                -Property Id,DisplayName,UserPrincipalName,OnPremisesSyncEnabled,OnPremisesImmutableId `
                -ErrorAction Stop
        }
        catch {
            continue
        }

        if ([string]::IsNullOrWhiteSpace($user.UserPrincipalName)) {
            continue
        }

        $isSynced = $false
        if ($user.OnPremisesSyncEnabled -eq $true -or $user.OnPremisesImmutableId) {
            $isSynced = $true
        }

        $rawUserRoles += [pscustomobject]@{
            UserId          = $user.Id
            DisplayName     = $user.DisplayName
            UPN             = $user.UserPrincipalName
            Role            = $role.DisplayName
            SyncedFromOnPrem= $isSynced
        }
    }
}

if (-not $rawUserRoles) {
    Write-Host "`nNo admin users found in the selected roles." -ForegroundColor Yellow
    return
}


# ------------------------------
# 4. Aggregate roles per user
# ------------------------------
$grouped = $rawUserRoles | Group-Object UPN

$adminList = @()

foreach ($g in $grouped) {
    $first = $g.Group | Select-Object -First 1
    $roles = $g.Group.Role | Sort-Object -Unique

    $adminList += [pscustomobject]@{
        Control               = "1.1.2 (L1)"
        DisplayName           = $first.DisplayName
        UPN                   = $first.UPN
        UserId                = $first.UserId
        Roles                 = $roles
        SyncedFromOnPrem      = ($g.Group | Where-Object { $_.SyncedFromOnPrem } | Measure-Object).Count -gt 0

        MFAEnabled            = $false
        MFA_Mechanisms        = @()
        MFA_HasAuthenticatorApp      = $false
        MFA_HasPasswordlessApp       = $false
        MFA_HasFIDO2                 = $false
        MFA_HasPhone                 = $false
        MFA_HasSoftwareOATH          = $false
        MFA_HasTAP                   = $false

        CA_MFARequired        = $false
        CA_MFAPolicies        = @()
    }
}


# ------------------------------
# 5. Load Conditional Access policies that require MFA
# ------------------------------
Write-Host "`nLoading Conditional Access policies that require MFA..." -ForegroundColor Cyan

$caPolicies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction SilentlyContinue | Where-Object {
    $_.State -eq "enabled" -and
    $_.GrantControls -and
    $_.GrantControls.BuiltInControls -contains "mfa"
}

if (-not $caPolicies) {
    Write-Host "No enabled Conditional Access policies explicitly requiring MFA were found." -ForegroundColor Yellow
}


# Cache for group memberships
$UserGroupCache = @{}

function Get-UserGroupIds {
    param(
        [string]$UserId
    )

    if ($UserGroupCache.ContainsKey($UserId)) {
        return $UserGroupCache[$UserId]
    }

    $groupIds = @()
    try {
        $groups = Get-MgUserMemberOf -UserId $UserId -All -ErrorAction SilentlyContinue
        foreach ($g in $groups) {
            if ($g.'@odata.type' -eq "#microsoft.graph.group") {
                $groupIds += $g.Id
            }
        }
    }
    catch {
        # ignore errors, return what we have (possibly empty)
    }

    $UserGroupCache[$UserId] = $groupIds
    return $groupIds
}


# ------------------------------
# 6. Build MFA profile per user
# ------------------------------
Write-Host "`nBuilding MFA profile for each admin account..." -ForegroundColor Cyan

foreach ($entry in $adminList) {

    $mfaMethods  = @()
    $mfaByCA     = $false
    $caPolicyHit = @()

    #
    # 6a. Authentication Methods
    #
    try {
        $methods = Get-MgUserAuthenticationMethod -UserId $entry.UserId -ErrorAction Stop

        foreach ($m in $methods) {
            # OData type is usually in AdditionalProperties["@odata.type"]
            $oType = $null
            if ($m.AdditionalProperties.ContainsKey("@odata.type")) {
                $oType = $m.AdditionalProperties["@odata.type"]
            }

            switch ($oType) {
                "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
                    $entry.MFA_HasAuthenticatorApp = $true
                    $mfaMethods += "AuthenticatorApp"
                }
                "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" {
                    $entry.MFA_HasPasswordlessApp = $true
                    $mfaMethods += "PasswordlessAuthenticatorApp"
                }
                "#microsoft.graph.fido2AuthenticationMethod" {
                    $entry.MFA_HasFIDO2 = $true
                    $mfaMethods += "FIDO2"
                }
                "#microsoft.graph.phoneAuthenticationMethod" {
                    $entry.MFA_HasPhone = $true
                    # Could be SMS/voice – both are valid MFA channels
                    $mfaMethods += "Phone"
                }
                "#microsoft.graph.softwareOathAuthenticationMethod" {
                    $entry.MFA_HasSoftwareOATH = $true
                    $mfaMethods += "SoftwareOATH"
                }
                "#microsoft.graph.temporaryAccessPassAuthenticationMethod" {
                    $entry.MFA_HasTAP = $true
                    $mfaMethods += "TemporaryAccessPass"
                }
                default {
                    if ($oType) {
                        $mfaMethods += $oType
                    }
                }
            }
        }
    }
    catch {
        # If we can't read methods, leave MFA flags as-is (all false)
    }

    #
    # 6b. Conditional Access: does any enabled CA policy requiring MFA apply?
    #
    if ($caPolicies -and $entry.UserId) {

        $userGroups = Get-UserGroupIds -UserId $entry.UserId

        foreach ($policy in $caPolicies) {
            $usersCond = $policy.Conditions.Users
            if (-not $usersCond) { continue }

            $includeUsers   = @($usersCond.IncludeUsers)
            $includeGroups  = @($usersCond.IncludeGroups)
            $excludeUsers   = @($usersCond.ExcludeUsers)
            $excludeGroups  = @($usersCond.ExcludeGroups)

            $isExcludedUser  = $excludeUsers -contains $entry.UserId
            $isExcludedGroup = ($excludeGroups | Where-Object { $userGroups -contains $_ }) -ne $null

            if ($isExcludedUser -or $isExcludedGroup) { continue }

            $includeAll      = $includeUsers -contains "All"
            $includedUser    = $includeUsers -contains $entry.UserId
            $includedGroup   = ($includeGroups | Where-Object { $userGroups -contains $_ }) -ne $null

            if ($includeAll -or $includedUser -or $includedGroup) {
                $mfaByCA = $true
                $caPolicyHit += $policy.DisplayName
            }
        }
    }

    #
    # 6c. Final MFA flags
    #
    $entry.MFA_Mechanisms = $mfaMethods | Sort-Object -Unique
    $entry.CA_MFARequired = $mfaByCA
    $entry.CA_MFAPolicies = $caPolicyHit | Sort-Object -Unique

    $entry.MFAEnabled = ($entry.MFA_HasAuthenticatorApp `
                         -or $entry.MFA_HasPasswordlessApp `
                         -or $entry.MFA_HasFIDO2 `
                         -or $entry.MFA_HasPhone `
                         -or $entry.MFA_HasSoftwareOATH `
                         -or $entry.MFA_HasTAP `
                         -or $entry.CA_MFARequired)
}


# ------------------------------
# 7. Final projection and output
# ------------------------------
$final = $adminList | Select-Object `
    Control,
    DisplayName,
    UPN,
    @{Name="Roles";Expression={($_.Roles -join ", ")}},
    SyncedFromOnPrem,
    MFAEnabled,
    @{Name="MFA_Mechanisms";Expression={($_.MFA_Mechanisms -join ", ")}},
    MFA_HasAuthenticatorApp,
    MFA_HasPasswordlessApp,
    MFA_HasFIDO2,
    MFA_HasPhone,
    MFA_HasSoftwareOATH,
    MFA_HasTAP,
    CA_MFARequired,
    @{Name="CA_MFAPolicies";Expression={($_.CA_MFAPolicies -join ", ")}}

Write-Host ""
Write-Host "===== CIS 1.1.2 – Administrative Accounts (Emergency Access MFA Profile) =====" -ForegroundColor Cyan
Write-Host "Total administrative users found: $($final.Count)"
Write-Host "A HUMAN must verify that exactly two appropriate break-glass accounts exist," -ForegroundColor Yellow
Write-Host "AND that their MFA behaviour matches your emergency-account standard." -ForegroundColor Yellow
Write-Host "===============================================================================" -ForegroundColor Cyan

$final | Sort-Object DisplayName | Format-Table -AutoSize


# ------------------------------
# 8. Export CSV if requested
# ------------------------------
if ($ReportPath) {
    $final | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nExported to: $ReportPath" -ForegroundColor Cyan
}

Write-Host "`nCIS 1.1.2 check complete.`n" -ForegroundColor Cyan
