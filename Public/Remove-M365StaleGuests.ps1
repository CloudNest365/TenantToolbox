function Remove-M365StaleGuests {
    <#
    .SYNOPSIS
        Finds and removes inactive guest accounts (with -WhatIf).
    .DESCRIPTION
        Lists guest accounts (userType = Guest) that have not signed in for a given number of
        days (or never) and removes them via Microsoft Graph. Every deletion goes through
        ShouldProcess: with -WhatIf NOTHING is deleted (dry run), only shown. Returns one
        object per (would-be) removed guest. Requires User.ReadWrite.All (write).
    .PARAMETER InactiveDays
        Threshold in days. Default: 90.
    .PARAMETER IncludeNeverSignedIn
        Also include guests that have never signed in.
    .PARAMETER WhatIf
        Dry run - only show what would be removed.
    .EXAMPLE
        Remove-M365StaleGuests -InactiveDays 180 -WhatIf
    .EXAMPLE
        Remove-M365StaleGuests -InactiveDays 180 -IncludeNeverSignedIn
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([pscustomobject])]
    param(
        [int]$InactiveDays = 90,
        [switch]$IncludeNeverSignedIn
    )

    Assert-TTGraph

    $cutoff = (Get-Date).AddDays(-$InactiveDays)
    Write-TTLog -Level INFO -Message "Searching for stale guests (threshold: $InactiveDays days)."

    $props = 'id,displayName,userPrincipalName,mail,createdDateTime,signInActivity'
    $guests = Get-MgUser -All -Filter "userType eq 'Guest'" -Property $props -ErrorAction Stop |
        Select-Object DisplayName, UserPrincipalName, Mail, CreatedDateTime, Id,
            @{ N = 'LastSignIn'; E = { $_.SignInActivity.LastSignInDateTime } }

    foreach ($g in $guests) {
        $last = $g.LastSignIn
        $never = $null -eq $last

        if ($never) {
            if (-not $IncludeNeverSignedIn) { continue }
            $daysInactive = $null
        }
        else {
            if ([datetime]$last -ge $cutoff) { continue }
            $daysInactive = [math]::Floor(((Get-Date) - [datetime]$last).TotalDays)
        }

        $removed = $false
        if ($PSCmdlet.ShouldProcess($g.UserPrincipalName, "Remove stale guest (last sign-in: $(if ($never) { 'never' } else { $last }))")) {
            try {
                Remove-MgUser -UserId $g.Id -ErrorAction Stop
                Write-TTLog -Level ACTION -Message "Removed stale guest '$($g.UserPrincipalName)' ($($g.Id))."
                $removed = $true
            }
            catch { Write-TTLog -Level WARN -Message "Could not remove '$($g.UserPrincipalName)': $_" }
        }

        [pscustomobject]@{
            DisplayName       = $g.DisplayName
            UserPrincipalName = $g.UserPrincipalName
            Mail              = $g.Mail
            LastSignIn        = if ($never) { $null } else { [datetime]$last }
            DaysInactive      = $daysInactive
            NeverSignedIn     = $never
            Removed           = $removed
            WhatIf            = [bool]$WhatIfPreference
            Id                = $g.Id
        }
    }
}
