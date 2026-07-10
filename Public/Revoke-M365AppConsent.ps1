function Revoke-M365AppConsent {
    <#
    .SYNOPSIS
        Revokes an enterprise app's delegated OAuth consent grants (remediation for the
        enterprise-app report).
    .DESCRIPTION
        Deletes the OAuth2 permission grants of a given app (by service-principal object id) via
        Microsoft Graph, removing its consented delegated permissions. Every deletion goes through
        ShouldProcess: with -WhatIf NOTHING is changed. Accepts pipeline input from
        Get-M365EnterpriseApp (ClientId). Requires DelegatedPermissionGrant.ReadWrite.All.
    .PARAMETER ClientId
        Service-principal (enterprise app) object id whose grants should be revoked.
    .PARAMETER WhatIf
        Dry run - only show what would be revoked.
    .EXAMPLE
        Get-M365EnterpriseApp -RiskyOnly | Revoke-M365AppConsent -WhatIf
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$ClientId
    )

    begin { Assert-TTGraph }

    process {
        # App display name (best effort)
        $appName = $ClientId
        try { $appName = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$ClientId`?`$select=displayName" -OutputType PSObject -ErrorAction Stop).displayName } catch { }

        $grants = @()
        try { $grants = @(Get-TTGraphCollection "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$filter=clientId eq '$ClientId'" -NoProgress) } catch { Write-TTLog -Level WARN -Message "Could not read grants for '$appName': $_" }

        $revoked = 0
        foreach ($g in $grants) {
            if ($PSCmdlet.ShouldProcess("$appName", "Revoke consent grant ($($g.scope))")) {
                try {
                    Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants/$($g.id)" -ErrorAction Stop | Out-Null
                    Write-TTLog -Level ACTION -Message "Revoked consent grant for '$appName' ($($g.id))."
                    $revoked++
                }
                catch { Write-TTLog -Level WARN -Message "Could not revoke grant for '$appName': $_" }
            }
        }

        [pscustomobject]@{ App = $appName; ClientId = $ClientId; GrantsFound = @($grants).Count; Revoked = $revoked; WhatIf = [bool]$WhatIfPreference }
    }
}
