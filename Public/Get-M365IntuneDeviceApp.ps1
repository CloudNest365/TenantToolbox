function Get-M365IntuneDeviceApp {
    <#
    .SYNOPSIS
        Returns the detected apps installed on a single Intune device (drilldown).
    .DESCRIPTION
        Resolves a device by name or id and lists its detected apps (name + version + size) via
        Microsoft Graph. Direct Graph calls, no extra submodule needed. Read-only.
    .PARAMETER DeviceName
        Device name to resolve (exact match).
    .PARAMETER DeviceId
        Managed device id (if you already have it).
    .EXAMPLE
        Get-M365IntuneDeviceApp -DeviceName 'DESKTOP-A19F'
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByName')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, ParameterSetName = 'ByName')]
        [string]$DeviceName,

        [Parameter(Mandatory, ParameterSetName = 'ById')]
        [string]$DeviceId
    )

    Assert-TTGraph

    if ($PSCmdlet.ParameterSetName -eq 'ByName') {
        $enc = [uri]::EscapeDataString($DeviceName)
        $found = Get-TTGraphCollection "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=deviceName eq '$enc'&`$select=id,deviceName"
        if (-not $found) { throw "Device '$DeviceName' not found." }
        $DeviceId = $found[0].id
    }

    Write-TTLog -Level INFO -Message "Reading detected apps for device $DeviceId ..."
    $apps = Get-TTGraphCollection "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$DeviceId/detectedApps?`$select=displayName,version,sizeInByte&`$top=100"

    foreach ($a in $apps) {
        [pscustomobject]@{
            DeviceId = $DeviceId
            App      = $a.displayName
            Version  = $a.version
            SizeMB   = if ($a.sizeInByte) { [math]::Round([double]$a.sizeInByte / 1MB, 1) } else { $null }
        }
    }
}
