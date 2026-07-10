function Export-M365IntuneAppDeploymentReport {
    <#
    .SYNOPSIS
        Generates an interactive HTML report of Intune app deployment status.
    .DESCRIPTION
        Builds on Get-M365IntuneAppDeployment and renders a self-contained HTML page: a KPI
        overview (apps, assigned, with failures, total failed installs) and a searchable,
        filterable, sortable table with assignment count and install/failed/not-installed
        device counts per app. Read-only.
    .PARAMETER Path
        Target path of the HTML file. Default: .\IntuneAppDeployment-Report.html
    .PARAMETER BrandName
        Branding. For CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER PassThru
        Also emit the deployment objects on the pipeline.
    .PARAMETER NoOpen
        Do not open the report automatically.
    .EXAMPLE
        Export-M365IntuneAppDeploymentReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'IntuneAppDeployment-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv,
        [switch]$Excel,
        [string]$DataPath,
        [switch]$NoHtml,
        [switch]$PassThru,
        [switch]$NoOpen
    )

    Assert-TTGraph
    $data = @(Get-M365IntuneAppDeployment)

    $total       = $data.Count
    $assigned    = @($data | Where-Object { $_.Assignments -gt 0 }).Count
    $withFail    = @($data | Where-Object { $_.Failed -gt 0 }).Count
    $totalFailed = ($data | Measure-Object Failed -Sum).Sum
    $genAt       = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $tenant      = (Get-MgContext).TenantId

    $rows = foreach ($a in ($data | Sort-Object Failed, @{E = { -$_.Assignments } }, App)) {
        $fAssigned = if ($a.Assignments -gt 0) { '1' } else { '0' }
        $fFail = if ($a.Failed -gt 0) { '1' } else { '0' }
        $failBadge = if ($a.Failed -gt 0) { "<span class='b b-bad'>$($a.Failed)</span>" } else { "<span class='muted'>0</span>" }
        $assignBadge = if ($a.Assignments -gt 0) { "<span class='b b-info'>$($a.Assignments)</span>" } else { "<span class='muted'>0</span>" }
        $failSort = ('{0:D7}' -f [int]$a.Failed)
        $searchAttr = TTEnc ("$($a.App) $($a.Type) $($a.Publisher)".ToLower())
        $nameAttr = TTEnc ([string]$a.App).ToLower()
        @"
      <tr class="item" data-f-assigned="$fAssigned" data-f-failures="$fFail" data-name="$nameAttr" data-s-failed="$failSort" data-search="$searchAttr">
        <td><div class="u"><b>$(TTEnc $a.App)</b><span class="upn">$(if ($a.Publisher) { TTEnc $a.Publisher } else { '' })</span></div></td>
        <td>$(TTEnc $a.Type)</td>
        <td>$assignBadge</td>
        <td>$($a.Installed)</td>
        <td data-s-failed="$failSort">$failBadge</td>
        <td>$($a.NotInstalled)</td>
      </tr>
"@
    }

    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Apps'; filter = 'all' }
        @{ n = $assigned; l = 'Assigned'; kind = 'info'; filter = 'assigned' }
        @{ n = $withFail; l = 'With failures'; kind = 'bad'; filter = 'failures' }
        @{ n = $totalFailed; l = 'Failed installs'; kind = 'bad' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search app, type or publisher ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'Assigned'; key = 'assigned' }, @{ label = 'With failures'; key = 'failures' }
    )
    $body = @"
    <div class="panel">
      <table class="tbl">
        <thead><tr><th data-sort="name">Application</th><th>Type</th><th>Assignments</th><th>Installed</th><th data-sort="failed">Failed</th><th>Not installed</th></tr></thead>
        <tbody>
$($rows -join "`n")
        </tbody>
      </table>
      <div class="empty" id="empty">No app matches the filters.</div>
    </div>
"@

    $sub = "Tenant $tenant &middot; generated $genAt &middot; $total apps, $totalFailed failed installs"
    $html = New-TTHtmlPage -Title 'Intune App Deployment' -Heading 'Intune App Deployment' -Sub $sub `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Intune app deployment report created: $Path ($total apps, $withFail with failures)."
    if ($withFail -gt 0) { Write-Host "  Note: $withFail app(s) have failed installs ($totalFailed devices)." -ForegroundColor Yellow }

    $flat = $data | Select-Object App, Type, Publisher, Assignments, Installed, Failed, NotInstalled, Pending
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Intune-App-Deployment-Report'
    if ($PassThru) { $data }
}
