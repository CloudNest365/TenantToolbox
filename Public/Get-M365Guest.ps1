function Get-M365Guest {
    <#
    .SYNOPSIS
        Returns guest (external) accounts with domain, state and last sign-in.
    .DESCRIPTION
        Reads guest users via Microsoft Graph and derives the external domain, invitation state and
        days since last sign-in. Read-only.
    .PARAMETER InactiveDays
        Mark guests stale if not signed in for this many days. Default: 90.
    .PARAMETER StaleOnly
        Return only stale guests.
    .EXAMPLE
        Get-M365Guest -StaleOnly
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([int]$InactiveDays = 90, [switch]$StaleOnly)

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading guest accounts ..."
    $cutoff = (Get-Date).AddDays(-$InactiveDays)

    $guests = Get-MgUser -All -Filter "userType eq 'Guest'" -Property 'id,displayName,userPrincipalName,mail,createdDateTime,externalUserState,signInActivity' -ErrorAction Stop

    foreach ($g in $guests) {
        $mail = if ($g.Mail) { $g.Mail } else { $g.UserPrincipalName }
        $domain = if ($mail -match '@') { ($mail -split '@')[-1] } else { '' }
        $domain = $domain -replace '#EXT#.*$', ''
        $last = $g.SignInActivity.LastSignInDateTime
        $never = $null -eq $last
        $days = if ($never) { $null } else { [math]::Floor(((Get-Date) - [datetime]$last).TotalDays) }
        $stale = $never -or ([datetime]$last -lt $cutoff)
        if ($StaleOnly -and -not $stale) { continue }

        [pscustomobject]@{
            DisplayName = $g.DisplayName
            Mail        = $g.Mail
            Domain      = $domain
            State       = $g.ExternalUserState
            Created     = $g.CreatedDateTime
            LastSignIn  = if ($never) { $null } else { [datetime]$last }
            DaysInactive = $days
            Stale       = $stale
            NeverSignedIn = $never
            Id          = $g.Id
        }
    }
}
