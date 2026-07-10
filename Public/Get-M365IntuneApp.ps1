function Get-M365IntuneApp {
    <#
    .SYNOPSIS
        Returns detected software (apps) across Intune-managed devices, with version and device count.
    .DESCRIPTION
        Reads the detected apps inventory via Microsoft Graph (deviceManagement/detectedApps) using
        direct Graph calls (no extra submodule needed). Each object is a distinct app + version with
        the number of devices it is installed on. Read-only. Data source for Export-M365IntuneAppReport.
    .PARAMETER MinDevices
        Only return apps installed on at least this many devices. Default: 1.
    .PARAMETER NameLike
        Optional wildcard filter on the app name (e.g. '*Chrome*').
    .EXAMPLE
        Get-M365IntuneApp -MinDevices 5 | Sort-Object Devices -Descending
    .EXAMPLE
        Get-M365IntuneApp -NameLike '*Acrobat*'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [int]$MinDevices = 1,
        [string]$NameLike
    )

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading Intune detected apps (software inventory) ..."

    $sel = 'displayName,version,sizeInByte,deviceCount,platform,publisher'
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/detectedApps?`$select=$sel&`$top=100"
    $apps = Get-TTGraphCollection $uri

    foreach ($a in $apps) {
        if ([int]$a.deviceCount -lt $MinDevices) { continue }
        if ($NameLike -and $a.displayName -notlike $NameLike) { continue }

        [pscustomobject]@{
            App          = $a.displayName
            Version      = $a.version
            Publisher    = $a.publisher
            Platform     = $a.platform
            Devices      = [int]$a.deviceCount
            SizeMB       = if ($a.sizeInByte) { [math]::Round([double]$a.sizeInByte / 1MB, 1) } else { $null }
        }
    }
}
