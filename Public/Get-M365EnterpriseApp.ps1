function Get-M365EnterpriseApp {
    <#
    .SYNOPSIS
        Returns enterprise apps and their consented (delegated) permissions, flagging risky scopes.
    .DESCRIPTION
        Reads OAuth2 permission grants via Microsoft Graph, aggregates the consented scopes per app,
        resolves the app name and marks risky scopes (e.g. Mail.ReadWrite, Directory.ReadWrite.All,
        full_access_as_user). Read-only. Uses Directory.Read.All.
    .PARAMETER RiskyOnly
        Return only apps that hold at least one risky scope.
    .EXAMPLE
        Get-M365EnterpriseApp -RiskyOnly
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([switch]$RiskyOnly)

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading OAuth2 permission grants (enterprise apps) ..."

    $risky = @(
        'Mail.ReadWrite', 'Mail.Send', 'MailboxSettings.ReadWrite', 'Files.ReadWrite.All',
        'Sites.ReadWrite.All', 'Directory.ReadWrite.All', 'User.ReadWrite.All', 'Group.ReadWrite.All',
        'full_access_as_user', 'Application.ReadWrite.All', 'RoleManagement.ReadWrite.Directory',
        'AppRoleAssignment.ReadWrite.All', 'Mail.Read', 'Files.Read.All'
    )

    try { $grants = Get-TTGraphCollection 'https://graph.microsoft.com/v1.0/oauth2PermissionGrants?$top=100' }
    catch { Write-Warning "Could not read oauth2PermissionGrants: $_"; return }

    $spCache = @{}
    function Resolve-Sp { param($Id)
        if ($spCache.ContainsKey($Id)) { return $spCache[$Id] }
        $n = $Id
        try { $n = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$Id`?`$select=displayName" -OutputType PSObject -ErrorAction Stop).displayName } catch { }
        $spCache[$Id] = $n; return $n
    }

    foreach ($g in ($grants | Group-Object clientId)) {
        $scopes = @($g.Group | ForEach-Object { ($_.scope -split '\s+') } | Where-Object { $_ } | Select-Object -Unique)
        $riskyScopes = @($scopes | Where-Object { $_ -in $risky })
        if ($RiskyOnly -and -not $riskyScopes) { continue }
        $tenantWide = @($g.Group | Where-Object { $_.consentType -eq 'AllPrincipals' }).Count -gt 0

        [pscustomobject]@{
            App         = Resolve-Sp $g.Name
            TenantWide  = $tenantWide
            ScopeCount  = $scopes.Count
            Scopes      = $scopes
            RiskyScopes = $riskyScopes
            ClientId    = $g.Name
        }
    }
}
