function Export-M365IntuneDeviceAppReport {
    <#
    .SYNOPSIS
        Generates an interactive HTML drilldown of the apps installed on a single device.
    .DESCRIPTION
        Builds on Get-M365IntuneDeviceApp and renders a self-contained HTML page for one device:
        a searchable, sortable table of its detected apps with version and size. Read-only.
    .PARAMETER DeviceName
        Device name to report on.
    .PARAMETER Path
        Target path of the HTML file. Default: .\IntuneDeviceApp-<DeviceName>.html
    .PARAMETER BrandName
        Branding. For CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER PassThru
        Also emit the app objects on the pipeline.
    .PARAMETER NoOpen
        Do not open the report automatically.
    .EXAMPLE
        Export-M365IntuneDeviceAppReport -DeviceName 'DESKTOP-A19F' -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$DeviceName,
        [string]$Path,
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
    $data = @(Get-M365IntuneDeviceApp -DeviceName $DeviceName | Sort-Object App)
    if (-not $Path) {
        $safe = ($DeviceName -replace '[^\w.-]', '_')
        $Path = Join-Path (Get-Location) "IntuneDeviceApp-$safe.html"
    }

    $total  = $data.Count
    $genAt  = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $tenant = (Get-MgContext).TenantId

    $rows = foreach ($a in $data) {
        $searchAttr = TTEnc ("$($a.App) $($a.Version)".ToLower())
        $nameAttr = TTEnc ([string]$a.App).ToLower()
        @"
      <tr class="item" data-name="$nameAttr" data-search="$searchAttr">
        <td><b>$(TTEnc $a.App)</b></td>
        <td><span class="chip">$(if ($a.Version) { TTEnc $a.Version } else { '–' })</span></td>
        <td>$(if ($null -ne $a.SizeMB) { "$($a.SizeMB) MB" } else { '<span class="muted">–</span>' })</td>
      </tr>
"@
    }

    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Installed apps'; filter = 'all' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search app or version ...' -Filters @( @{ label = 'All'; key = 'all' } )
    $body = @"
    <div class="panel">
      <table class="tbl">
        <thead><tr><th data-sort="name">Application</th><th>Version</th><th>Size</th></tr></thead>
        <tbody>
$($rows -join "`n")
        </tbody>
      </table>
      <div class="empty" id="empty">No app matches the search.</div>
    </div>
"@

    $sub = "Device <b>$(TTEnc $DeviceName)</b> &middot; Tenant $tenant &middot; generated $genAt &middot; $total apps"
    $html = New-TTHtmlPage -Title "Device Apps - $DeviceName" -Heading "Installed Apps - $DeviceName" -Sub $sub `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Device app drilldown created: $Path ($total apps on '$DeviceName')."

    $flat = $data | Select-Object App, Version, SizeMB
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Intune-Device-App-Report'
    if ($PassThru) { $data }
}
