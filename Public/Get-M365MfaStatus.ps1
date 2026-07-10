function Get-M365MfaStatus {
    <#
    .SYNOPSIS
        Returns the MFA / authentication registration status of all users.
    .DESCRIPTION
        Uses the Graph 'userRegistrationDetails' report and returns structured objects:
        whether MFA is registered/capable, admin yes/no, default method, registered
        methods, etc. Read-only. Ideal as the data source for Export-M365MfaReport or
        for direct processing (Export-Csv, Where-Object ...).
    .PARAMETER UnregisteredOnly
        Return only users WITHOUT registered MFA.
    .PARAMETER AdminsOnly
        Return only users with privileged roles (isAdmin).
    .PARAMETER IncludeGuests
        Also include guest accounts (default: members only).
    .EXAMPLE
        Get-M365MfaStatus -UnregisteredOnly | Export-Csv without-mfa.csv -NoTypeInformation
    .EXAMPLE
        Get-M365MfaStatus -AdminsOnly | Where-Object { -not $_.MfaRegistered }
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [switch]$UnregisteredOnly,
        [switch]$AdminsOnly,
        [switch]$IncludeGuests
    )

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading MFA registration status (userRegistrationDetails) ..."

    $details = Get-MgReportAuthenticationMethodUserRegistrationDetail -All -ErrorAction Stop

    foreach ($d in $details) {
        if (-not $IncludeGuests -and $d.UserType -eq 'guest') { continue }
        if ($AdminsOnly -and -not $d.IsAdmin) { continue }
        if ($UnregisteredOnly -and $d.IsMfaRegistered) { continue }

        [pscustomobject]@{
            DisplayName         = $d.UserDisplayName
            UserPrincipalName   = $d.UserPrincipalName
            UserType            = $d.UserType
            IsAdmin             = [bool]$d.IsAdmin
            MfaRegistered       = [bool]$d.IsMfaRegistered
            MfaCapable          = [bool]$d.IsMfaCapable
            SsprRegistered      = [bool]$d.IsSsprRegistered
            PasswordlessCapable = [bool]$d.IsPasswordlessCapable
            DefaultMethod       = $d.DefaultMfaMethod
            Methods             = @($d.MethodsRegistered)
            LastUpdated         = $d.LastUpdatedDateTime
            Id                  = $d.Id
        }
    }
}
