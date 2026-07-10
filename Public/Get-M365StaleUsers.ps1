function Get-M365StaleUsers {
    <#
    .SYNOPSIS
        Finds users who have not signed in for X days (or never).
    .DESCRIPTION
        Reads the last interactive sign-in date of all users via Microsoft Graph and
        returns structured objects - pipeline-friendly via Export-Csv / ConvertTo-Html.
        Read-only, changes nothing.
    .PARAMETER InactiveDays
        Threshold in days. Default: 90.
    .PARAMETER IncludeNeverSignedIn
        Also include accounts that have never signed in.
    .PARAMETER IncludeDisabled
        Also include already-disabled accounts (default: only enabled).
    .EXAMPLE
        Get-M365StaleUsers -InactiveDays 90 | Export-Csv stale.csv -NoTypeInformation
    .EXAMPLE
        Get-M365StaleUsers -InactiveDays 180 -IncludeNeverSignedIn
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [int]$InactiveDays = 90,

        [switch]$IncludeNeverSignedIn,

        [switch]$IncludeDisabled
    )

    Assert-TTGraph

    $cutoff = (Get-Date).AddDays(-$InactiveDays)
    Write-TTLog -Level INFO -Message "Searching for inactive users (threshold: $InactiveDays days, before $($cutoff.ToString('yyyy-MM-dd')))."

    $props = 'id,displayName,userPrincipalName,mail,accountEnabled,signInActivity,department,jobTitle'
    $users = Get-MgUser -All -Property $props -ErrorAction Stop |
        Select-Object DisplayName, UserPrincipalName, Mail, AccountEnabled, Department, JobTitle, Id,
            @{ N = 'LastSignIn'; E = { $_.SignInActivity.LastSignInDateTime } }

    foreach ($u in $users) {
        if (-not $IncludeDisabled -and -not $u.AccountEnabled) { continue }

        $last       = $u.LastSignIn
        $neverLogin = $null -eq $last

        if ($neverLogin) {
            if (-not $IncludeNeverSignedIn) { continue }
            $daysInactive = $null
        }
        else {
            if ([datetime]$last -ge $cutoff) { continue }
            $daysInactive = [math]::Floor(((Get-Date) - [datetime]$last).TotalDays)
        }

        [pscustomobject]@{
            DisplayName       = $u.DisplayName
            UserPrincipalName = $u.UserPrincipalName
            Mail              = $u.Mail
            Department        = $u.Department
            JobTitle          = $u.JobTitle
            AccountEnabled    = $u.AccountEnabled
            LastSignIn        = if ($neverLogin) { $null } else { [datetime]$last }
            DaysInactive      = $daysInactive
            NeverSignedIn     = $neverLogin
            Id                = $u.Id
        }
    }
}
