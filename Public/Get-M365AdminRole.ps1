function Get-M365AdminRole {
    <#
    .SYNOPSIS
        Returns directory role assignments (who holds which admin role), cross-referenced with MFA.
    .DESCRIPTION
        Reads activated directory roles and their members via Microsoft Graph and joins each member
        with their MFA registration status. Read-only. Uses Directory.Read.All + AuditLog.Read.All.
    .PARAMETER PrivilegedOnly
        Return only assignments for privileged/high-impact roles.
    .EXAMPLE
        Get-M365AdminRole -PrivilegedOnly | Where-Object { -not $_.MfaRegistered }
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([switch]$PrivilegedOnly)

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading directory role assignments ..."

    $privileged = @(
        'Global Administrator', 'Privileged Role Administrator', 'Privileged Authentication Administrator',
        'Security Administrator', 'Conditional Access Administrator', 'Exchange Administrator',
        'SharePoint Administrator', 'User Administrator', 'Application Administrator',
        'Cloud Application Administrator', 'Hybrid Identity Administrator', 'Intune Administrator',
        'Authentication Administrator', 'Helpdesk Administrator'
    )

    $mfaMap = @{}
    try { Get-M365MfaStatus | ForEach-Object { if ($_.UserPrincipalName) { $mfaMap[$_.UserPrincipalName.ToLower()] = [bool]$_.MfaRegistered } } } catch { }

    $roles = Get-TTGraphCollection 'https://graph.microsoft.com/v1.0/directoryRoles?$expand=members'

    foreach ($role in $roles) {
        $isPriv = $privileged -contains $role.displayName
        if ($PrivilegedOnly -and -not $isPriv) { continue }
        foreach ($m in @($role.members)) {
            $t = [string]$m.'@odata.type'
            $type = if ($t -match 'user') { 'user' } elseif ($t -match 'group') { 'group' } elseif ($t -match 'servicePrincipal') { 'sp' } else { 'other' }
            $upn = [string]$m.userPrincipalName
            $mfa = if ($type -eq 'user' -and $upn -and $mfaMap.ContainsKey($upn.ToLower())) { $mfaMap[$upn.ToLower()] } else { $null }
            [pscustomobject]@{
                Role          = $role.displayName
                Privileged    = $isPriv
                Member        = [string]$m.displayName
                UPN           = $upn
                Type          = $type
                MfaRegistered = $mfa
            }
        }
    }
}
