# ============================================================
# STEP 1 - Install the Graph module (only need to do this once)
# ============================================================
Install-Module Microsoft.Graph -Scope CurrentUser


# ============================================================
# STEP 2 - Connect to Graph
# ============================================================
Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All"


# ============================================================
# STEP 3 - Set your inactive threshold
# ============================================================
$DaysInactive = 60
$CutoffDate   = (Get-Date).AddDays(-$DaysInactive)


# ============================================================
# STEP 4 - Pull all licensed, enabled users with sign-in data
# ============================================================
$Users = Get-MgUser -All `
    -Filter "assignedLicenses/`$count ne 0 and accountEnabled eq true" `
    -CountVariable TotalCount `
    -ConsistencyLevel eventual `
    -Property "id, displayName, userPrincipalName, signInActivity, assignedLicenses"

$Users.Count  # sanity check — how many users came back?


# ============================================================
# STEP 5 - Filter to inactive users only
# ============================================================
$InactiveUsers = $Users | Where-Object {
    $lastSignIn = $_.SignInActivity?.LastSignInDateTime
    ($null -eq $lastSignIn) -or ([datetime]$lastSignIn -lt $CutoffDate)
}

$InactiveUsers.Count  # how many are inactive?


# ============================================================
# STEP 6 - View the results
# ============================================================
$InactiveUsers | Select-Object DisplayName, UserPrincipalName,
    @{ Name = "LastSignIn"; Expression = { $_.SignInActivity?.LastSignInDateTime ?? "Never" } } |
    Sort-Object LastSignIn |
    Format-Table -AutoSize


# ============================================================
# STEP 7 - Export to CSV (optional but recommended before touching anything)
# ============================================================
$InactiveUsers | Select-Object DisplayName, UserPrincipalName,
    @{ Name = "LastSignIn"; Expression = { $_.SignInActivity?.LastSignInDateTime ?? "Never" } } |
    Export-Csv -Path ".\InactiveUsers_$(Get-Date -Format 'yyyyMMdd').csv" -NoTypeInformation
