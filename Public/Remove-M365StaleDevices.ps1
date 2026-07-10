function Remove-M365StaleDevices {
    <#
    .SYNOPSIS
        Finds and removes stale Intune devices (no sync for X days), with -WhatIf.
    .DESCRIPTION
        Lists managed devices that have not synced for a given number of days and deletes them
        from Intune via Microsoft Graph. Every deletion goes through ShouldProcess: with -WhatIf
        NOTHING is deleted (dry run), only shown. Returns one object per (would-be) removed device.
        Requires DeviceManagementManagedDevices.ReadWrite.All (write).
    .PARAMETER StaleDays
        A device counts as stale if it has not synced for this many days. Default: 90.
    .PARAMETER WhatIf
        Dry run - only show what would be removed.
    .EXAMPLE
        Remove-M365StaleDevices -StaleDays 120 -WhatIf
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([pscustomobject])]
    param(
        [int]$StaleDays = 90
    )

    Assert-TTGraph
    $stale = @(Get-M365IntuneDevice -StaleDays $StaleDays -StaleOnly)
    Write-TTLog -Level INFO -Message "Found $($stale.Count) stale device(s) (no sync for $StaleDays days)."

    foreach ($d in $stale) {
        $removed = $false
        if ($PSCmdlet.ShouldProcess($d.DeviceName, "Delete stale device (last sync: $(if ($d.LastSync) { $d.LastSync.ToString('yyyy-MM-dd') } else { 'never' }))")) {
            try {
                Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$($d.DeviceId)" -ErrorAction Stop | Out-Null
                Write-TTLog -Level ACTION -Message "Deleted stale device '$($d.DeviceName)' ($($d.DeviceId))."
                $removed = $true
            }
            catch { Write-TTLog -Level WARN -Message "Could not delete '$($d.DeviceName)': $_" }
        }

        [pscustomobject]@{
            DeviceName    = $d.DeviceName
            User          = $d.User
            OS            = $d.OS
            LastSync      = $d.LastSync
            DaysSinceSync = $d.DaysSinceSync
            Removed       = $removed
            WhatIf        = [bool]$WhatIfPreference
            DeviceId      = $d.DeviceId
        }
    }
}
