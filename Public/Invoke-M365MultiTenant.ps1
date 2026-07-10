function Invoke-M365MultiTenant {
    <#
    .SYNOPSIS
        Runs a scriptblock of read cmdlets across all tenants set via Set-M365MultiTenant.
    .DESCRIPTION
        Connects to each configured tenant in turn, runs the given scriptblock, and returns the
        combined results - each object tagged with a 'Tenant' property. Errors on one tenant do not
        stop the others. Use only READ cmdlets here; state-changing cmdlets across many tenants are
        intentionally not encouraged.
    .PARAMETER ScriptBlock
        The cmdlets to run per tenant, e.g. { Get-M365MfaStatus }.
    .PARAMETER UseDeviceCode
        For interactive mode, use device code (default). Ignored in app-only mode.
    .EXAMPLE
        Invoke-M365MultiTenant { Get-M365StaleUsers -InactiveDays 90 } | Export-Csv all-stale.csv -NoTypeInformation
    .EXAMPLE
        Invoke-M365MultiTenant { Get-M365MfaStatus | Where-Object { -not $_.MfaRegistered } }
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, Position = 0)][scriptblock]$ScriptBlock,
        [switch]$UseDeviceCode
    )

    $cfg = $script:TTMultiTenant
    if (-not $cfg) { throw "No multi-tenant context set. Run Set-M365MultiTenant first." }

    foreach ($tid in $cfg.TenantIds) {
        $name = if ($cfg.Names -and $cfg.Names.ContainsKey($tid)) { $cfg.Names[$tid] } else { $tid }
        Write-Host "== Tenant: $name ($tid) ==" -ForegroundColor Cyan

        try {
            if ($cfg.AppOnly) {
                Connect-TenantToolbox -TenantId $tid -ClientId $cfg.ClientId -CertificateThumbprint $cfg.CertificateThumbprint
            }
            else {
                Connect-TenantToolbox -TenantId $tid -UseDeviceCode:([bool]$UseDeviceCode -or $true)
            }
        }
        catch { Write-Warning "[$name] sign-in failed: $_"; continue }

        if (-not $script:TTConnected) { Write-Warning "[$name] not connected - skipped."; continue }

        try {
            $out = & $ScriptBlock
            foreach ($o in $out) {
                if ($null -ne $o -and $o.PSObject -and $o.GetType().Name -eq 'PSCustomObject') {
                    $o | Add-Member -NotePropertyName Tenant -NotePropertyValue $name -Force -PassThru
                }
                else { $o }
            }
        }
        catch { Write-Warning "[$name] query failed: $_" }
    }
}
