function Get-M365IntuneDevice {
    <#
    .SYNOPSIS
        Returns Intune-managed devices with compliance, sync and encryption info.
    .DESCRIPTION
        Reads managed devices via Microsoft Graph (deviceManagement/managedDevices) using direct
        Graph calls (no extra submodule needed) and returns structured objects. Read-only. Ideal
        as the data source for Export-M365IntuneDeviceReport or for direct processing.
    .PARAMETER StaleDays
        A device counts as stale if it has not synced for this many days. Default: 30.
    .PARAMETER NonCompliantOnly
        Return only devices that are not compliant.
    .PARAMETER StaleOnly
        Return only stale devices (no sync within StaleDays).
    .EXAMPLE
        Get-M365IntuneDevice -NonCompliantOnly | Export-Csv noncompliant.csv -NoTypeInformation
    .EXAMPLE
        Get-M365IntuneDevice -StaleDays 60 -StaleOnly
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [int]$StaleDays = 30,
        [switch]$NonCompliantOnly,
        [switch]$StaleOnly
    )

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading Intune managed devices ..."

    $sel = 'id,deviceName,userDisplayName,userPrincipalName,operatingSystem,osVersion,complianceState,managedDeviceOwnerType,lastSyncDateTime,isEncrypted,model,manufacturer'
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=$sel&`$top=100"
    $devices = Get-TTGraphCollection $uri

    $now = Get-Date
    foreach ($d in $devices) {
        $last = $d.lastSyncDateTime
        $daysSince = if ($last) { [math]::Floor(($now - [datetime]$last).TotalDays) } else { $null }
        $stale = ($null -ne $daysSince) -and ($daysSince -gt $StaleDays)

        if ($NonCompliantOnly -and $d.complianceState -eq 'compliant') { continue }
        if ($StaleOnly -and -not $stale) { continue }

        $owner = switch ([string]$d.managedDeviceOwnerType) {
            'company'  { 'Corporate' }
            'personal' { 'Personal' }
            default    { if ($d.managedDeviceOwnerType) { [string]$d.managedDeviceOwnerType } else { 'Unknown' } }
        }

        [pscustomobject]@{
            DeviceId      = $d.id
            DeviceName    = $d.deviceName
            User          = $d.userDisplayName
            UPN           = $d.userPrincipalName
            OS            = $d.operatingSystem
            OSVersion     = $d.osVersion
            Compliance    = $d.complianceState
            Owner         = $owner
            LastSync      = if ($last) { [datetime]$last } else { $null }
            DaysSinceSync = $daysSince
            Stale         = $stale
            Encrypted     = [bool]$d.isEncrypted
            Model         = $d.model
            Manufacturer  = $d.manufacturer
        }
    }
}
