function Export-M365AutopilotReport {
    <#
    .SYNOPSIS
        Generates an HTML report of Windows Autopilot devices and profile assignment.
    .DESCRIPTION
        Reads windowsAutopilotDeviceIdentities via Graph and shows serial, model, group tag,
        profile assignment status and enrollment state. Read-only. Requires
        DeviceManagementServiceConfig.Read.All.
    .PARAMETER Path
        Target path of the HTML file. Default: .\Autopilot-Report.html
    .EXAMPLE
        Export-M365AutopilotReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'Autopilot-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading Autopilot devices ..."
    try { $ap = Get-TTGraphCollection "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?`$top=100" }
    catch {
        if ("$_" -match 'Forbidden|403') { Write-Warning "Access denied. Needs DeviceManagementServiceConfig.Read.All. Reconnect: Connect-TenantToolbox -UseDeviceCode" }
        else { Write-Warning "Could not read Autopilot devices: $_" }
        $ap = @()
    }

    $data = foreach ($d in $ap) {
        [pscustomobject]@{
            Serial = $d.serialNumber; Model = $d.model; Manufacturer = $d.manufacturer; GroupTag = $d.groupTag
            ProfileStatus = $d.deploymentProfileAssignmentStatus; Enrollment = $d.enrollmentState; LastContact = $d.lastContactedDateTime
        }
    }
    $data = @($data)
    $total = $data.Count
    $unassigned = @($data | Where-Object { $_.ProfileStatus -match 'notAssigned|pending' }).Count
    $notEnrolled = @($data | Where-Object { $_.Enrollment -notmatch 'enrolled' }).Count
    $genAt = Get-Date -Format 'yyyy-MM-dd HH:mm'; $tenant = (Get-MgContext).TenantId

    $rows = foreach ($d in ($data | Sort-Object Model, Serial)) {
        $fUnassigned = if ($d.ProfileStatus -match 'notAssigned|pending') { '1' } else { '0' }
        $fNotEnrolled = if ($d.Enrollment -notmatch 'enrolled') { '1' } else { '0' }
        $ps = if ($d.ProfileStatus -match 'assigned' -and $d.ProfileStatus -notmatch 'not') { "<span class='b b-ok'>$(TTEnc $d.ProfileStatus)</span>" } else { "<span class='b b-warn'>$(TTEnc $d.ProfileStatus)</span>" }
        $searchAttr = TTEnc ("$($d.Serial) $($d.Model) $($d.GroupTag)".ToLower()); $nameAttr = TTEnc ([string]$d.Serial).ToLower()
        @"
      <tr class="item" data-f-unassigned="$fUnassigned" data-f-notenrolled="$fNotEnrolled" data-name="$nameAttr" data-search="$searchAttr">
        <td><div class="u"><b>$(TTEnc $d.Serial)</b><span class="upn">$(TTEnc $d.Manufacturer) $(TTEnc $d.Model)</span></div></td>
        <td>$(if ($d.GroupTag) { TTEnc $d.GroupTag } else { '<span class="muted">&#8211;</span>' })</td>
        <td>$ps</td><td>$(TTEnc $d.Enrollment)</td>
      </tr>
"@
    }
    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Autopilot devices'; filter = 'all' }
        @{ n = $unassigned; l = 'No profile'; kind = 'warn'; filter = 'unassigned' }
        @{ n = $notEnrolled; l = 'Not enrolled'; kind = 'bad'; filter = 'notenrolled' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search serial, model or tag ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'No profile'; key = 'unassigned' }, @{ label = 'Not enrolled'; key = 'notenrolled' }
    )
    $body = @"
    <div class="panel"><table class="tbl"><thead><tr><th data-sort="name">Serial</th><th>Group tag</th><th>Profile</th><th>Enrollment</th></tr></thead>
      <tbody>
$($rows -join "`n")
      </tbody></table><div class="empty" id="empty">No device matches the filters.</div></div>
"@
    $html = New-TTHtmlPage -Title 'Autopilot' -Heading 'Windows Autopilot' -Sub "Tenant $tenant &middot; generated $genAt &middot; $total devices" `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Autopilot report created: $Path ($total devices, $unassigned without profile)."
    $flat = $data | Select-Object Serial, Model, Manufacturer, GroupTag, ProfileStatus, Enrollment
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Autopilot-Report'
    if ($PassThru) { $data }
}
