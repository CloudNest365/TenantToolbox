function Export-M365IntuneAppReport {
    <#
    .SYNOPSIS
        Generates an interactive HTML software inventory report from Intune detected apps.
    .DESCRIPTION
        Builds on Get-M365IntuneApp and renders a self-contained HTML page: a KPI overview
        (apps, installs, unique titles, widely deployed) and a searchable, filterable, sortable
        table of software with version, publisher, platform and device count. Read-only.
    .PARAMETER Path
        Target path of the HTML file. Default: .\IntuneApp-Report.html
    .PARAMETER MinDevices
        Only include apps installed on at least this many devices. Default: 1.
    .PARAMETER WideThreshold
        Device count from which an app counts as "widely deployed". Default: 10.
    .PARAMETER BrandName
        Branding. For CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER PassThru
        Also emit the app objects on the pipeline.
    .PARAMETER NoOpen
        Do not open the report automatically.
    .EXAMPLE
        Export-M365IntuneAppReport -MinDevices 3 -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'IntuneApp-Report.html'),
        [int]$MinDevices = 1,
        [int]$WideThreshold = 10,
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
    $data = @(Get-M365IntuneApp -MinDevices $MinDevices)

    function Get-Os { param($os)
        $o = [string]$os
        if ($o -match 'windows') { 'windows' } elseif ($o -match 'ios|ipad') { 'ios' }
        elseif ($o -match 'android') { 'android' } elseif ($o -match 'mac') { 'macos' } else { 'other' }
    }

    # KPIs
    $total    = $data.Count
    $installs = ($data | Measure-Object Devices -Sum).Sum
    $titles   = @($data | Select-Object -ExpandProperty App -Unique).Count
    $wide     = @($data | Where-Object { $_.Devices -ge $WideThreshold }).Count
    $genAt    = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $tenant   = (Get-MgContext).TenantId

    $rows = foreach ($a in ($data | Sort-Object Devices -Descending)) {
        $osKey = Get-Os $a.Platform
        $fWide = if ($a.Devices -ge $WideThreshold) { '1' } else { '0' }
        $devSort = ('{0:D7}' -f [int]$a.Devices)
        $searchAttr = TTEnc ("$($a.App) $($a.Version) $($a.Publisher)".ToLower())
        $nameAttr = TTEnc ([string]$a.App).ToLower()
        @"
      <tr class="item" data-f-wide="$fWide" data-f-$osKey="1" data-name="$nameAttr" data-s-devices="$devSort" data-search="$searchAttr">
        <td><div class="u"><b>$(TTEnc $a.App)</b><span class="upn">$(if ($a.Publisher) { TTEnc $a.Publisher } else { '' })</span></div></td>
        <td><span class="chip">$(if ($a.Version) { TTEnc $a.Version } else { '–' })</span></td>
        <td>$(if ($a.Platform) { TTEnc $a.Platform } else { '<span class="muted">–</span>' })</td>
        <td data-s-devices="$devSort"><b>$($a.Devices)</b></td>
        <td>$(if ($null -ne $a.SizeMB) { "$($a.SizeMB) MB" } else { '<span class="muted">–</span>' })</td>
      </tr>
"@
    }

    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'App versions'; filter = 'all' }
        @{ n = $titles; l = 'Unique titles'; kind = 'info' }
        @{ n = $installs; l = 'Total installs'; kind = 'ok' }
        @{ n = $wide; l = "On $WideThreshold+ devices"; kind = 'warn'; filter = 'wide' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search app, version or publisher ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = "On $WideThreshold+ devices"; key = 'wide' },
        @{ label = 'Windows'; key = 'windows' }, @{ label = 'iOS'; key = 'ios' }, @{ label = 'Android'; key = 'android' }, @{ label = 'macOS'; key = 'macos' }
    )
    $body = @"
    <div class="panel">
      <table class="tbl">
        <thead><tr><th data-sort="name">Application</th><th>Version</th><th>Platform</th><th data-sort="devices">Devices</th><th>Size</th></tr></thead>
        <tbody>
$($rows -join "`n")
        </tbody>
      </table>
      <div class="empty" id="empty">No application matches the filters.</div>
    </div>
"@

    $sub = "Tenant $tenant &middot; generated $genAt &middot; $total app versions on $installs installs"
    $html = New-TTHtmlPage -Title 'Intune Software Inventory' -Heading 'Intune Software Inventory' -Sub $sub `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Intune app report created: $Path ($total app versions, $titles titles, $installs installs)."

    $flat = $data | Select-Object App, Version, Publisher, Platform, Devices, SizeMB
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Intune-App-Report'
    if ($PassThru) { $data }
}
