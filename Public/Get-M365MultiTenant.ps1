function Get-M365MultiTenant {
    <#
    .SYNOPSIS
        Shows the currently configured multi-tenant context (set via Set-M365MultiTenant).
    .EXAMPLE
        Get-M365MultiTenant
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    if (-not $script:TTMultiTenant) {
        Write-Warning "No multi-tenant context set. Run Set-M365MultiTenant first."
        return
    }
    $script:TTMultiTenant
}
