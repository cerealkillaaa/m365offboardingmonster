[CmdletBinding()]
param(
    [string]$ConfigPath = ".\config.json",
    [switch]$WhatIfMode
)

$ErrorActionPreference = "Stop"

# ==================================================
# GLOBALS
# ==================================================

$Script:LogRoot = Join-Path $PSScriptRoot "Logs"
$Script:RunTime = Get-Date -Format "yyyyMMdd-HHmmss"
$Script:LogFile = Join-Path $Script:LogRoot "Offboarding-$Script:RunTime.log"

# ==================================================
# BANNER
# ==================================================

function Show-Banner {

    Clear-Host

    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "           M365 OFFBOARDING MONSTER v1.0.0" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
}

# ==================================================
# LOGGING
# ==================================================

function Ensure-LogFolder {

    if (-not (Test-Path $Script:LogRoot)) {
        New-Item -ItemType Directory -Path $Script:LogRoot | Out-Null
    }
}

function Write-Log {

    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $line = "[$timestamp] [$Level] $Message"

    Add-Content -Path $Script:LogFile -Value $line

    $color = switch ($Level) {
        "OK" { "Green" }
        "WARN" { "Yellow" }
        "ERROR" { "Red" }
        default { "White" }
    }

    Write-Host $line -ForegroundColor $color
}

# ==================================================
# CONFIG
# ==================================================

function Load-Config {

    param([string]$Path)

    if (-not (Test-Path $Path)) {
        throw "Config file not found: $Path"
    }

    return Get-Content $Path -Raw | ConvertFrom-Json
}

# ==================================================
# MODULES
# ==================================================

function Import-RequiredModules {

    Write-Log "Loading required modules..."

    Import-Module ActiveDirectory -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
}

# ==================================================
# CONNECTIONS
# ==================================================

function Connect-Services {

    Write-Log "Connecting to Microsoft Graph..."

    Connect-MgGraph `
        -Scopes `
        "User.ReadWrite.All",
        "Directory.ReadWrite.All",
        "Organization.Read.All" `
        -NoWelcome

    Write-Log "Connecting to Exchange Online..."

    Connect-ExchangeOnline -ShowBanner:$false
}

# ==================================================
# HELPERS
# ==================================================

function Confirm-Execution {

    param([string]$Prompt)

    Write-Host ""
    Write-Host $Prompt -ForegroundColor Yellow
    Write-Host ""

    $response = Read-Host "Type YES to continue"

    return $response -eq "YES"
}

function Get-ADUserSafe {

    param([string]$Identity)

    try {

        return Get-ADUser `
            -Identity $Identity `
            -Properties `
            DisplayName,
            DistinguishedName,
            UserPrincipalName,
            MemberOf,
            mail

    }
    catch {

        throw "Unable to locate AD user: $Identity"
    }
}

function Invoke-Step {

    param(
        [string]$Name,
        [scriptblock]$Action
    )

    try {

        Write-Log "START: $Name"

        if ($WhatIfMode) {

            Write-Log "WHATIF: Skipping execution for $Name" "WARN"
        }
        else {

            & $Action
        }

        Write-Log "SUCCESS: $Name" "OK"
    }
    catch {

        Write-Log "FAILED: $Name - $($_.Exception.Message)" "ERROR"
    }
}

# ==================================================
# AD ACTIONS
# ==================================================

function Disable-ADUserAccount {

    param(
        [Microsoft.ActiveDirectory.Management.ADUser]$User
    )

    Disable-ADAccount -Identity $User.DistinguishedName
}

function Move-ToDisabledOU {

    param(
        [Microsoft.ActiveDirectory.Management.ADUser]$User,
        [string]$DisabledOU
    )

    Move-ADObject `
        -Identity $User.DistinguishedName `
        -TargetPath $DisabledOU
}

function Remove-FromPrivilegedGroups {

    param(
        [Microsoft.ActiveDirectory.Management.ADUser]$User,
        [array]$ProtectedGroups
    )

    foreach ($groupDN in $User.MemberOf) {

        $group = Get-ADGroup $groupDN

        if ($ProtectedGroups -contains $group.Name) {

            Remove-ADGroupMember `
                -Identity $group `
                -Members $User `
                -Confirm:$false

            Write-Log "Removed from privileged group: $($group.Name)"
        }
    }
}

# ==================================================
# M365 / ENTRA ACTIONS
# ==================================================

function Disable-EntraSignIn {

    param([string]$UPN)

    Update-MgUser `
        -UserId $UPN `
        -AccountEnabled:$false
}

function Revoke-Sessions {

    param([string]$UPN)

    Revoke-MgUserSignInSession `
        -UserId $UPN | Out-Null
}

function Remove-M365Licenses {

    param([string]$UPN)

    $user = Get-MgUser -UserId $UPN

    $licenses = Get-MgUserLicenseDetail -UserId $user.Id

    if ($licenses.Count -eq 0) {

        Write-Log "No licenses assigned."
        return
    }

    $removeLicenses = @()

    foreach ($license in $licenses) {

        $removeLicenses += $license.SkuId
    }

    Set-MgUserLicense `
        -UserId $user.Id `
        -AddLicenses @() `
        -RemoveLicenses $removeLicenses
}

# ==================================================
# EXCHANGE ACTIONS
# ==================================================

function Convert-ToSharedMailbox {

    param([string]$UPN)

    Set-Mailbox `
        -Identity $UPN `
        -Type Shared
}

function Hide-FromGAL {

    param([string]$UPN)

    Set-Mailbox `
        -Identity $UPN `
        -HiddenFromAddressListsEnabled $true
}

function Configure-MailForwarding {

    param(
        [string]$UPN,
        [string]$ManagerEmail
    )

    Set-Mailbox `
        -Identity $UPN `
        -ForwardingSMTPAddress $ManagerEmail `
        -DeliverToMailboxAndForward $true
}

function Grant-ManagerMailboxAccess {

    param(
        [string]$UPN,
        [string]$ManagerEmail
    )

    Add-MailboxPermission `
        -Identity $UPN `
        -User $ManagerEmail `
        -AccessRights FullAccess `
        -InheritanceType All `
        -AutoMapping:$true
}

# ==================================================
# REPORT
# ==================================================

function Export-OffboardingSummary {

    param(
        [string]$Employee,
        [string]$Manager
    )

    $summary = @"
==========================================
OFFBOARDING SUMMARY
==========================================

Employee: $Employee
Manager:  $Manager
Timestamp: $(Get-Date)

Actions:
- AD Account Disabled
- AD Account Moved
- Privileged Groups Removed
- Entra Sign-In Disabled
- M365 Sessions Revoked
- Exchange Mailbox Converted
- Mailbox Hidden from GAL
- Mail Forwarding Enabled
- Licenses Removed

==========================================
"@

    $summaryFile = Join-Path $Script:LogRoot "Summary-$Script:RunTime.txt"

    Set-Content `
        -Path $summaryFile `
        -Value $summary

    Write-Log "Summary report created: $summaryFile"
}

# ==================================================
# MAIN WORKFLOW
# ==================================================

function Start-Offboarding {

    Show-Banner

    Ensure-LogFolder

    Write-Log "Loading configuration..."

    $config = Load-Config -Path $ConfigPath

    Import-RequiredModules

    Connect-Services

    Write-Host ""

    $employee = Read-Host "Employee username"
    $managerEmail = Read-Host "Manager email"

    $user = Get-ADUserSafe -Identity $employee

    Write-Host ""
    Write-Host "==================================================" -ForegroundColor Yellow
    Write-Host "OFFBOARDING PREVIEW"
    Write-Host "==================================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Employee: $($user.DisplayName)"
    Write-Host "UPN:      $($user.UserPrincipalName)"
    Write-Host "Manager:  $managerEmail"
    Write-Host ""
    Write-Host "Actions:"
    Write-Host " - Disable AD account"
    Write-Host " - Move account to disabled OU"
    Write-Host " - Remove privileged access"
    Write-Host " - Disable Entra sign-in"
    Write-Host " - Revoke active sessions"
    Write-Host " - Convert mailbox to shared"
    Write-Host " - Hide from GAL"
    Write-Host " - Forward mail"
    Write-Host " - Remove licenses"
    Write-Host ""

    if (-not (Confirm-Execution "Continue with offboarding?")) {

        Write-Log "Operation cancelled by operator." "WARN"
        return
    }

    Invoke-Step "Disable AD Account" {
        Disable-ADUserAccount -User $user
    }

    Invoke-Step "Move AD User To Disabled OU" {
        Move-ToDisabledOU `
            -User $user `
            -DisabledOU $config.DisabledOU
    }

    Invoke-Step "Remove Privileged Group Memberships" {
        Remove-FromPrivilegedGroups `
            -User $user `
            -ProtectedGroups $config.PrivilegedGroups
    }

    Invoke-Step "Disable Entra Sign-In" {
        Disable-EntraSignIn `
            -UPN $user.UserPrincipalName
    }

    Invoke-Step "Revoke Sessions" {
        Revoke-Sessions `
            -UPN $user.UserPrincipalName
    }

    Invoke-Step "Convert Mailbox To Shared" {
        Convert-ToSharedMailbox `
            -UPN $user.UserPrincipalName
    }

    Invoke-Step "Hide Mailbox From GAL" {
        Hide-FromGAL `
            -UPN $user.UserPrincipalName
    }

    Invoke-Step "Configure Mail Forwarding" {
        Configure-MailForwarding `
            -UPN $user.UserPrincipalName `
            -ManagerEmail $managerEmail
    }

    Invoke-Step "Grant Manager Mailbox Access" {
        Grant-ManagerMailboxAccess `
            -UPN $user.UserPrincipalName `
            -ManagerEmail $managerEmail
    }

    Invoke-Step "Remove M365 Licenses" {
        Remove-M365Licenses `
            -UPN $user.UserPrincipalName
    }

    Export-OffboardingSummary `
        -Employee $employee `
        -Manager $managerEmail

    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Green
    Write-Host "OFFBOARDING COMPLETE"
    Write-Host "======================================================" -ForegroundColor Green
    Write-Host ""
}

Start-Offboarding
