function Get-M365Group {
    <#
    .SYNOPSIS
        Returns groups with type, owner count and orphan flag.
    .DESCRIPTION
        Reads groups via Microsoft Graph (with owners expanded) and derives the group type
        (Microsoft 365 / Security / Distribution), owner count and an orphan flag (Microsoft 365
        group with no owner). Read-only.
    .PARAMETER OrphanedOnly
        Return only orphaned groups (Microsoft 365 groups with no owner).
    .EXAMPLE
        Get-M365Group -OrphanedOnly
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([switch]$OrphanedOnly)

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading groups ..."

    $groups = Get-TTGraphCollection 'https://graph.microsoft.com/v1.0/groups?$select=id,displayName,mailEnabled,securityEnabled,groupTypes,visibility,createdDateTime&$expand=owners($select=id)&$top=100'

    foreach ($g in $groups) {
        $isUnified = @($g.groupTypes) -contains 'Unified'
        $type = if ($isUnified) { 'Microsoft 365' }
                elseif ($g.securityEnabled -and -not $g.mailEnabled) { 'Security' }
                elseif ($g.mailEnabled -and -not $g.securityEnabled) { 'Distribution' }
                elseif ($g.mailEnabled -and $g.securityEnabled) { 'Mail-enabled security' }
                else { 'Other' }
        $owners = @($g.owners).Count
        $orphaned = $isUnified -and $owners -eq 0
        if ($OrphanedOnly -and -not $orphaned) { continue }

        [pscustomobject]@{
            DisplayName = $g.displayName
            Type        = $type
            Owners      = $owners
            Orphaned    = $orphaned
            Visibility  = $g.visibility
            Created     = $g.createdDateTime
            Id          = $g.id
        }
    }
}
