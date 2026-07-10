function Set-M365MultiTenant {
    <#
    .SYNOPSIS
        Defines a set of tenants to run read cmdlets against with Invoke-M365MultiTenant.
    .DESCRIPTION
        Stores a list of tenant ids (and optional certificate/app details) in the module session.
        Afterwards, Invoke-M365MultiTenant { <cmdlets> } runs any read cmdlets across all of them
        and tags each result with a 'Tenant' column.

        Two auth modes:
          - App-only (recommended for MSP / unattended): pass -ClientId and -CertificateThumbprint
            of a multi-tenant app that has admin-consent in every customer tenant.
          - Interactive: omit them; Invoke-M365MultiTenant then signs in per tenant via device code.
    .PARAMETER TenantId
        One or more tenant ids or default domains (e.g. contoso.onmicrosoft.com).
    .PARAMETER ClientId
        App (client) id of a multi-tenant app for app-only certificate auth.
    .PARAMETER CertificateThumbprint
        Certificate thumbprint for app-only auth.
    .PARAMETER Name
        Optional hashtable mapping a tenant id to a friendly display name.
    .EXAMPLE
        Set-M365MultiTenant -TenantId 'contoso.onmicrosoft.com','fabrikam.onmicrosoft.com'
    .EXAMPLE
        Set-M365MultiTenant -TenantId $ids -ClientId <appid> -CertificateThumbprint <thumb> `
            -Name @{ 'contoso.onmicrosoft.com' = 'Contoso' }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint,
        [hashtable]$Name
    )

    $appOnly = [bool]($ClientId -and $CertificateThumbprint)
    $script:TTMultiTenant = [pscustomobject]@{
        TenantIds             = @($TenantId)
        ClientId              = $ClientId
        CertificateThumbprint = $CertificateThumbprint
        Names                 = $Name
        AppOnly               = $appOnly
    }

    $mode = if ($appOnly) { 'app-only (certificate)' } else { 'interactive (device code per tenant)' }
    Write-Host "Multi-tenant context set: $($TenantId.Count) tenant(s), auth: $mode." -ForegroundColor Green
    Write-Host "Run:  Invoke-M365MultiTenant { Get-M365MfaStatus }" -ForegroundColor DarkGray
}
