function Export-M365LicenseReport {
    <#
    .SYNOPSIS
        Generates an interactive HTML report of license (SKU) assignment.
    .DESCRIPTION
        Builds on Get-M365License and renders a self-contained HTML page: a KPI overview (SKUs,
        assigned seats, available seats, near-full SKUs) and a searchable, sortable table with
        total / assigned / available seats and usage per SKU. Governance view, not billing. Read-only.
    .PARAMETER Path
        Target path of the HTML file. Default: .\License-Report.html
    .PARAMETER BrandName
        Branding. For CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER PassThru
        Also emit the objects on the pipeline.
    .PARAMETER NoOpen
        Do not open the report automatically.
    .EXAMPLE
        Export-M365LicenseReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'License-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    Assert-TTGraph
    $data = @(Get-M365License)

    $skus     = $data.Count
    $assigned = ($data | Measure-Object Assigned -Sum).Sum
    $avail    = ($data | Measure-Object Available -Sum).Sum
    $nearFull = @($data | Where-Object { $_.Total -gt 0 -and $_.UsagePct -ge 90 }).Count
    $genAt    = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $tenant   = (Get-MgContext).TenantId

    $rows = foreach ($s in ($data | Sort-Object Assigned -Descending)) {
        $fFull = if ($s.Total -gt 0 -and $s.UsagePct -ge 90) { '1' } else { '0' }
        $fUnused = if ($s.Assigned -eq 0) { '1' } else { '0' }
        $pctColor = if ($s.UsagePct -ge 100) { 'b-bad' } elseif ($s.UsagePct -ge 90) { 'b-warn' } else { 'b-ok' }
        $availTxt = if ($s.Available -lt 0) { "<span class='b b-bad'>$($s.Available)</span>" } else { "$($s.Available)" }
        $assignSort = ('{0:D7}' -f [int]$s.Assigned)
        $searchAttr = TTEnc ("$($s.License) $($s.SkuPart)".ToLower())
        $nameAttr = TTEnc ([string]$s.License).ToLower()
        @"
      <tr class="item" data-f-full="$fFull" data-f-unused="$fUnused" data-name="$nameAttr" data-s-assigned="$assignSort" data-search="$searchAttr">
        <td><div class="u"><b>$(TTEnc $s.License)</b><span class="upn">$(TTEnc $s.SkuPart)</span></div></td>
        <td>$($s.Total)</td>
        <td data-s-assigned="$assignSort"><b>$($s.Assigned)</b></td>
        <td>$availTxt</td>
        <td><span class="b $pctColor">$($s.UsagePct)%</span></td>
      </tr>
"@
    }

    $kpiHtml = New-TTKpis @(
        @{ n = $skus; l = 'SKUs'; filter = 'all' }
        @{ n = $assigned; l = 'Assigned seats'; kind = 'info' }
        @{ n = $avail; l = 'Available seats'; kind = 'ok' }
        @{ n = $nearFull; l = 'Near full (>=90%)'; kind = 'warn'; filter = 'full' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search license ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'Near full'; key = 'full' }, @{ label = 'Unused'; key = 'unused' }
    )
    $body = @"
    <div class="panel">
      <table class="tbl">
        <thead><tr><th data-sort="name">License</th><th>Total</th><th data-sort="assigned">Assigned</th><th>Available</th><th>Usage</th></tr></thead>
        <tbody>
$($rows -join "`n")
        </tbody>
      </table>
      <div class="empty" id="empty">No license matches the filters.</div>
    </div>
"@

    $sub = "Tenant $tenant &middot; generated $genAt &middot; $skus SKUs, $assigned seats assigned"
    $html = New-TTHtmlPage -Title 'License Assignment' -Heading 'License Assignment' -Sub $sub `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "License report created: $Path ($skus SKUs, $assigned assigned, $avail available)."

    $flat = $data | Select-Object License, SkuPart, Total, Assigned, Available, UsagePct
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'License-Report'
    if ($PassThru) { $data }
}
