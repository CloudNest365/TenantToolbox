function Get-M365RiskyUser {
    <#
    .SYNOPSIS
        Returns users flagged by Entra Identity Protection (risky users).
    .DESCRIPTION
        Reads identityProtection/riskyUsers via Microsoft Graph and returns risk level and state
        per user. Read-only. Requires Entra ID P2 and the IdentityRiskyUser.Read.All scope.
    .PARAMETER AtRiskOnly
        Return only users that are currently at risk or confirmed compromised.
    .EXAMPLE
        Get-M365RiskyUser -AtRiskOnly
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([switch]$AtRiskOnly)

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading risky users (Identity Protection) ..."
    try {
        $risky = Get-TTGraphCollection 'https://graph.microsoft.com/v1.0/identityProtection/riskyUsers?$top=100'
    }
    catch {
        if ("$_" -match 'Forbidden|403') {
            Write-Warning "Access denied. Needs Entra ID P2 and the 'IdentityRiskyUser.Read.All' scope. Reconnect: Connect-TenantToolbox -UseDeviceCode"
        }
        else { Write-Warning "Could not read risky users: $_" }
        return
    }

    foreach ($u in $risky) {
        if ($AtRiskOnly -and $u.riskState -notin 'atRisk', 'confirmedCompromised') { continue }
        [pscustomobject]@{
            User        = $u.userDisplayName
            UPN         = $u.userPrincipalName
            RiskLevel   = $u.riskLevel
            RiskState   = $u.riskState
            LastUpdated = $u.riskLastUpdatedDateTime
            Id          = $u.id
        }
    }
}
