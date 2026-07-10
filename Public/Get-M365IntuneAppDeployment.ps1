function Get-M365IntuneAppDeployment {
    <#
    .SYNOPSIS
        Returns Intune managed apps with their assignment and install status.
    .DESCRIPTION
        Reads mobileApps via Microsoft Graph (with assignments) and, for assigned apps, the
        install summary (installed / failed / not installed device counts). Direct Graph calls,
        no extra submodule needed. Read-only. Data source for Export-M365IntuneAppDeploymentReport.
    .PARAMETER AssignedOnly
        Return only apps that have at least one assignment.
    .PARAMETER WithFailuresOnly
        Return only apps that have failed installs.
    .EXAMPLE
        Get-M365IntuneAppDeployment -WithFailuresOnly
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [switch]$AssignedOnly,
        [switch]$WithFailuresOnly
    )

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading Intune app deployments ..."

    try {
        $apps = Get-TTGraphCollection "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps?`$expand=assignments&`$top=100"
    }
    catch {
        if ("$_" -match 'Forbidden|403') {
            Write-Warning "Access denied. This report needs the 'DeviceManagementApps.Read.All' scope. Reconnect: Connect-TenantToolbox -UseDeviceCode"
        }
        else { Write-Warning "Could not read mobileApps: $_" }
        return
    }

    foreach ($a in $apps) {
        $assignCount = @($a.assignments).Count
        if ($AssignedOnly -and $assignCount -eq 0) { continue }

        $type = ([string]$a.'@odata.type') -replace '#microsoft\.graph\.', ''

        $installed = 0; $failed = 0; $notInstalled = 0; $pending = 0
        if ($assignCount -gt 0) {
            try {
                $s = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps/$($a.id)/installSummary" -OutputType PSObject -ErrorAction Stop
                $installed    = [int]$s.installedDeviceCount
                $failed       = [int]$s.failedDeviceCount
                $notInstalled = [int]$s.notInstalledDeviceCount
                $pending      = [int]$s.pendingInstallDeviceCount
            }
            catch { }
        }

        if ($WithFailuresOnly -and $failed -eq 0) { continue }

        [pscustomobject]@{
            App          = $a.displayName
            Type         = $type
            Publisher    = $a.publisher
            Assignments  = $assignCount
            Installed    = $installed
            Failed       = $failed
            NotInstalled = $notInstalled
            Pending      = $pending
            AppId        = $a.id
        }
    }
}
