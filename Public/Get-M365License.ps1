function Get-M365License {
    <#
    .SYNOPSIS
        Returns license (SKU) assignment info: total, assigned and available seats.
    .DESCRIPTION
        Reads subscribedSkus via Microsoft Graph and returns per SKU the enabled (purchased),
        consumed (assigned) and available seats plus a usage percentage. Governance view, not
        billing. Read-only.
    .EXAMPLE
        Get-M365License | Sort-Object Available
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading subscribed SKUs (licenses) ..."

    $friendly = @{
        'ENTERPRISEPREMIUM' = 'Office 365 E5'; 'ENTERPRISEPACK' = 'Office 365 E3'
        'SPE_E3' = 'Microsoft 365 E3'; 'SPE_E5' = 'Microsoft 365 E5'; 'SPB' = 'Microsoft 365 Business Premium'
        'O365_BUSINESS_PREMIUM' = 'M365 Business Standard'; 'O365_BUSINESS_ESSENTIALS' = 'M365 Business Basic'
        'EMS' = 'EMS E3'; 'EMSPREMIUM' = 'EMS E5'; 'AAD_PREMIUM' = 'Entra ID P1'; 'AAD_PREMIUM_P2' = 'Entra ID P2'
        'FLOW_FREE' = 'Power Automate Free'; 'POWER_BI_STANDARD' = 'Power BI (free)'; 'POWER_BI_PRO' = 'Power BI Pro'
        'TEAMS_EXPLORATORY' = 'Teams Exploratory'; 'WINDOWS_STORE' = 'Windows Store'; 'MCOEV' = 'Teams Phone'
        'DEFENDER_ENDPOINT_P1' = 'Defender for Endpoint P1'; 'INTUNE_A' = 'Intune Plan 1'
    }

    $skus = Get-TTGraphCollection 'https://graph.microsoft.com/v1.0/subscribedSkus'

    foreach ($s in $skus) {
        $part = [string]$s.skuPartNumber
        $total = [int]$s.prepaidUnits.enabled
        $assigned = [int]$s.consumedUnits
        $avail = $total - $assigned
        $pct = if ($total -gt 0) { [math]::Round(100 * $assigned / $total) } else { 0 }
        [pscustomobject]@{
            License   = if ($friendly.ContainsKey($part)) { $friendly[$part] } else { $part }
            SkuPart   = $part
            Total     = $total
            Assigned  = $assigned
            Available = $avail
            UsagePct  = $pct
        }
    }
}
