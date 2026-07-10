function Get-M365MfaStatus {
    <#
    .SYNOPSIS
        Liefert den MFA-/Authentifizierungs-Registrierungsstatus aller Benutzer.
    .DESCRIPTION
        Nutzt den Graph-Report 'userRegistrationDetails' und gibt strukturierte Objekte
        zurueck: ob MFA registriert/faehig, Admin ja/nein, Standardmethode, registrierte
        Methoden usw. Reines Lesen. Ideal als Datenbasis fuer Export-M365MfaReport oder
        zum direkten Weiterverarbeiten (Export-Csv, Where-Object ...).
    .PARAMETER UnregisteredOnly
        Nur Benutzer OHNE registrierte MFA zurueckgeben.
    .PARAMETER AdminsOnly
        Nur Benutzer mit privilegierten Rollen (isAdmin) zurueckgeben.
    .PARAMETER IncludeGuests
        Auch Gastkonten aufnehmen (Standard: nur Mitglieder).
    .EXAMPLE
        Get-M365MfaStatus -UnregisteredOnly | Export-Csv ohne-mfa.csv -NoTypeInformation
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
    Write-TTLog -Level INFO -Message "Lese MFA-Registrierungsstatus (userRegistrationDetails) ..."

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
