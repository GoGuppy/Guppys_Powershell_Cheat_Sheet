<#
.SYNOPSIS
    Finds inactive licensed users (default: 60+ days), disables their accounts,
    and emails their manager requesting permission to archive and remove licensing.

.DESCRIPTION
    Uses Microsoft Graph PowerShell SDK.
    Requires the following Graph scopes:
        - User.ReadWrite.All       (disable accounts)
        - AuditLog.Read.All        (read sign-in activity)
        - Mail.Send                (send email as FromAddress)
        - Directory.Read.All       (read manager info)

    Run with -WhatIf (default) to preview changes without touching anything.
    Set -WhatIf:$false to go live.

.PARAMETER DaysInactive
    Number of days since last sign-in to consider a user inactive. Default: 60.

.PARAMETER FromAddress
    The licensed mailbox the notification email is sent FROM (e.g. helpdesk@contoso.com).

.PARAMETER WhatIf
    If $true (default), no changes are made — output shows what WOULD happen.
    Set to $false to actually disable accounts and send emails.

.PARAMETER LogPath
    Path to output CSV log. Default: .\InactiveUsers_<date>.csv

.EXAMPLE
    # Preview only — no changes
    .\Invoke-InactiveUserLockdown.ps1 -FromAddress "helpdesk@contoso.com"

.EXAMPLE
    # Go live
    .\Invoke-InactiveUserLockdown.ps1 -FromAddress "helpdesk@contoso.com" -WhatIf:$false

.EXAMPLE
    # Custom threshold, live run, custom log path
    .\Invoke-InactiveUserLockdown.ps1 -FromAddress "helpdesk@contoso.com" -DaysInactive 90 -WhatIf:$false -LogPath "C:\Logs\InactiveUsers.csv"

.NOTES
    Exam relevance: AZ-104 (Entra ID / Identity), MD-102 (Endpoint/User lifecycle)
    Requires: Microsoft.Graph PowerShell SDK
    Install:  Install-Module Microsoft.Graph -Scope CurrentUser
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [int]    $DaysInactive = 60,
    [string] $FromAddress  = "",
    [bool]   $WhatIf       = $true,
    [string] $LogPath      = ".\InactiveUsers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

#region ── Preflight ────────────────────────────────────────────────────────────

if (-not $FromAddress) {
    Write-Error "You must specify -FromAddress (the mailbox emails are sent from)."
    exit 1
}

# Check Graph module
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    Write-Error "Microsoft.Graph module not found. Run: Install-Module Microsoft.Graph -Scope CurrentUser"
    exit 1
}

Write-Host "`n=== Inactive User Lockdown ===" -ForegroundColor Cyan
Write-Host "Threshold : $DaysInactive days inactive"
Write-Host "From      : $FromAddress"
Write-Host "WhatIf    : $WhatIf"
Write-Host "Log       : $LogPath`n"

if ($WhatIf) {
    Write-Host "[WHATIF MODE] No changes will be made. Use -WhatIf:`$false to go live.`n" -ForegroundColor Yellow
}

#endregion

#region ── Connect ──────────────────────────────────────────────────────────────

$Scopes = @(
    "User.ReadWrite.All",
    "AuditLog.Read.All",
    "Mail.Send",
    "Directory.Read.All"
)

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes $Scopes -NoWelcome

#endregion

#region ── Get Inactive Licensed Users ─────────────────────────────────────────

$CutoffDate = (Get-Date).AddDays(-$DaysInactive)
$Results    = [System.Collections.Generic.List[PSCustomObject]]::new()

Write-Host "Fetching licensed users with sign-in activity..." -ForegroundColor Cyan

# Pull all users with licenses + sign-in activity in one call
# signInActivity requires AuditLog.Read.All
$Users = Get-MgUser -All `
    -Filter "assignedLicenses/`$count ne 0 and accountEnabled eq true" `
    -CountVariable UserCount `
    -ConsistencyLevel eventual `
    -Property "id,displayName,userPrincipalName,assignedLicenses,signInActivity,accountEnabled,mail"

Write-Host "Found $($Users.Count) licensed, enabled users. Checking sign-in activity...`n"

foreach ($User in $Users) {

    $LastSignIn = $User.SignInActivity?.LastSignInDateTime

    # Treat never-signed-in as inactive (no sign-in on record)
    $IsInactive = if ($null -eq $LastSignIn) {
        $true
    } else {
        [datetime]$LastSignIn -lt $CutoffDate
    }

    if (-not $IsInactive) { continue }

    $LastSignInDisplay = if ($null -eq $LastSignIn) { "Never" } else { $LastSignIn }

    # Get manager
    $Manager     = $null
    $ManagerMail = $null
    $ManagerName = $null
    try {
        $Manager     = Get-MgUserManager -UserId $User.Id -ErrorAction Stop
        $ManagerFull = Get-MgUser -UserId $Manager.Id -Property "displayName,mail" -ErrorAction Stop
        $ManagerMail = $ManagerFull.Mail
        $ManagerName = $ManagerFull.DisplayName
    } catch {
        $ManagerMail = $null
        $ManagerName = "No Manager Found"
    }

    Write-Host "INACTIVE: $($User.DisplayName) ($($User.UserPrincipalName))" -ForegroundColor Yellow
    Write-Host "  Last Sign-In : $LastSignInDisplay"
    Write-Host "  Manager      : $ManagerName ($ManagerMail)"

    # ── Disable Account ──────────────────────────────────────────────────────
    if (-not $WhatIf) {
        try {
            Update-MgUser -UserId $User.Id -AccountEnabled:$false
            Write-Host "  [DISABLED]   Account locked." -ForegroundColor Red
            $DisabledStatus = "Disabled"
        } catch {
            Write-Warning "  Failed to disable $($User.UserPrincipalName): $_"
            $DisabledStatus = "ERROR: $_"
        }
    } else {
        Write-Host "  [WHATIF]     Would disable account." -ForegroundColor DarkYellow
        $DisabledStatus = "WhatIf - Not disabled"
    }

    # ── Send Manager Email ────────────────────────────────────────────────────
    $EmailStatus = "No manager / not sent"

    if ($ManagerMail) {
        $Subject = "Action Required: $($User.DisplayName) account locked due to inactivity"

        $Body = @"
Hi $ManagerName,

This is an automated notification from IT.

<b>$($User.DisplayName)</b> ($($User.UserPrincipalName)) has been detected as inactive for 60+ days.
Their last recorded sign-in was: <b>$LastSignInDisplay</b>

As a result, the account has been <b>locked</b> pending your review.

<b>Please reply to confirm one of the following:</b>
<ul>
  <li>✅ Archive the account and remove licensing (user is no longer active)</li>
  <li>🔄 Re-enable the account (user still needs access)</li>
</ul>

If we do not hear back within 14 days, the account will be scheduled for archival.

Thanks,
IT Administration
"@

        $Message = @{
            subject      = $Subject
            body         = @{
                contentType = "HTML"
                content     = $Body
            }
            toRecipients = @(
                @{ emailAddress = @{ address = $ManagerMail } }
            )
        }

        if (-not $WhatIf) {
            try {
                Send-MgUserMail -UserId $FromAddress -Message $Message -SaveToSentItems:$true
                Write-Host "  [EMAILED]    Manager notified at $ManagerMail" -ForegroundColor Green
                $EmailStatus = "Sent to $ManagerMail"
            } catch {
                Write-Warning "  Failed to send email to $ManagerMail`: $_"
                $EmailStatus = "ERROR: $_"
            }
        } else {
            Write-Host "  [WHATIF]     Would email manager at $ManagerMail" -ForegroundColor DarkYellow
            $EmailStatus = "WhatIf - Would send to $ManagerMail"
        }
    }

    Write-Host ""

    # ── Log Entry ─────────────────────────────────────────────────────────────
    $Results.Add([PSCustomObject]@{
        DisplayName       = $User.DisplayName
        UPN               = $User.UserPrincipalName
        LastSignIn        = $LastSignInDisplay
        ManagerName       = $ManagerName
        ManagerEmail      = $ManagerMail
        AccountDisabled   = $DisabledStatus
        ManagerEmailSent  = $EmailStatus
        RunTimestamp      = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    })
}

#endregion

#region ── Output ───────────────────────────────────────────────────────────────

if ($Results.Count -eq 0) {
    Write-Host "No inactive licensed users found." -ForegroundColor Green
} else {
    Write-Host "=== Summary ===" -ForegroundColor Cyan
    Write-Host "Total inactive users found : $($Results.Count)"
    Write-Host "Log saved to               : $LogPath`n"
    $Results | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
    $Results | Format-Table -AutoSize
}

Disconnect-MgGraph | Out-Null
Write-Host "Done. Disconnected from Graph." -ForegroundColor Cyan

#endregion
